# db.py — MongoDB helper for DS Codification Intelligence System
from pymongo import MongoClient, DESCENDING
from datetime import datetime
from urllib.parse import quote_plus
import os

_raw_uri = os.environ.get("MONGO_URI", "mongodb+srv://defstdncell2026:DSCell@123@cluster0.d5czr7o.mongodb.net/?appName=Cluster0")

def _encode_mongo_uri(uri: str) -> str:
    """
    If the URI is in the standard Atlas format with user:password@,
    this re-encodes only the username and password portion using
    RFC 3986 percent-encoding so special characters don't break parsing.
    If it's already encoded or localhost, it's returned as-is.
    """
    try:
        if "@" in uri and ("mongodb+srv://" in uri or "mongodb://" in uri):
            scheme_end = uri.index("://") + 3          # after "mongodb+srv://"
            at_pos     = uri.rindex("@")               # last @ separates creds from host
            creds      = uri[scheme_end:at_pos]        # "user:password"
            rest       = uri[at_pos:]                  # "@cluster.../..."
            scheme     = uri[:scheme_end]              # "mongodb+srv://"
            if ":" in creds:
                user, password = creds.split(":", 1)
                encoded = f"{scheme}{quote_plus(user)}:{quote_plus(password)}{rest}"
                return encoded
    except Exception:
        pass
    return uri

MONGO_URI = _encode_mongo_uri(_raw_uri)
DB_NAME   = os.environ.get("MONGO_DB",  "ds_codification")

_client = None
_db     = None

def get_db():
    global _client, _db
    if _db is None:
        _client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=5000)
        _db     = _client[DB_NAME]
        _ensure_indexes()
    return _db

def _ensure_indexes():
    db = _client[DB_NAME]
    db.uploads.create_index([("uploaded_at", DESCENDING)])
    db.uploads.create_index([("filename", 1)])
    db.reports.create_index([("generated_at", DESCENDING)])
    db.reports.create_index([("source_upload_id", 1)])
    db.reports.create_index([("period_year", 1), ("period_month", 1)])

# ── Uploads ──────────────────────────────────────────────────────────────────
def save_upload(filename: str, display_name: str, rows: int,
                stats: dict, dpsu_list: list) -> str:
    """
    Persist metadata of an uploaded dataset.
    Returns the inserted document _id as string.
    """
    db = get_db()
    doc = {
        "filename":     filename,
        "display_name": display_name,
        "rows":         rows,
        "uploaded_at":  datetime.utcnow(),
        "stats": {
            "total":              stats.get("total", 0),
            "forwarded":          stats.get("forwarded", 0),
            "nsn_allotted":       stats.get("nsn_allotted", 0),
            "returned":           stats.get("returned", 0),
            "pending":            stats.get("pending", 0),
            "avg_mrc":            stats.get("avg_mrc", 0),
            "avg_processing_days":stats.get("avg_processing_days", 0),
            "by_dpsu":            stats.get("by_dpsu", {}),
            "by_ncb":             stats.get("by_ncb", {}),
            "by_equipment":       stats.get("by_equipment", {}),
        },
        "dpsu_list": dpsu_list,
    }
    result = db.uploads.insert_one(doc)
    return str(result.inserted_id)


def get_uploads(limit: int = 50) -> list:
    """Return recent uploads (metadata only, no raw Excel bytes)."""
    db   = get_db()
    docs = db.uploads.find({}, {"_id": 1, "filename": 1, "display_name": 1,
                                "rows": 1, "uploaded_at": 1, "stats": 1,
                                "dpsu_list": 1}) \
                     .sort("uploaded_at", DESCENDING).limit(limit)
    out = []
    for d in docs:
        d["_id"] = str(d["_id"])
        d["uploaded_at"] = d["uploaded_at"].strftime("%d %b %Y, %H:%M UTC")
        out.append(d)
    return out


# ── Reports ──────────────────────────────────────────────────────────────────
def save_report(excel_filename: str, heading: str, period_month: int,
                period_year: int, source_filename: str, display_name: str,
                totals: dict, rows: list, source_upload_id: str = None) -> str:
    """
    Persist full report data so quarterly / yearly aggregation is possible later.
    `rows` is the list of row dicts returned by generate_report() flattened.
    Returns inserted _id as string.
    """
    db = get_db()
    doc = {
        "excel_filename":   excel_filename,
        "heading":          heading,
        "period_month":     period_month,       # int 1-12
        "period_year":      period_year,        # int e.g. 2026
        "source_filename":  source_filename,
        "display_name":     display_name,
        "source_upload_id": source_upload_id,   # string ObjectId of upload doc
        "generated_at":     datetime.utcnow(),
        "totals": {
            "total_codified": totals.get("codified", 0),
            "fwd_dca":        totals.get("fwd", 0),
            "nsn_allotted":   totals.get("nsn", 0),
            "returned":       totals.get("returned", 0),
        },
        # Full row-level detail — enables quarterly/yearly drill-down later
        "rows": rows,   # list of {dpsu, equipment, total_codified, fwd_dca, nsn, returned}
    }
    result = db.reports.insert_one(doc)
    return str(result.inserted_id)


def get_reports(limit: int = 100, year: int = None,
                month: int = None, dpsu: str = None) -> list:
    """
    Flexible report query.
    Pass year/month/dpsu to filter; omit for all recent reports.
    """
    db     = get_db()
    query  = {}
    if year:  query["period_year"]  = year
    if month: query["period_month"] = month
    if dpsu:  query["rows.dpsu"]    = dpsu

    docs = db.reports.find(query, {"rows": 0}) \
                     .sort("generated_at", DESCENDING).limit(limit)
    out = []
    for d in docs:
        d["_id"]          = str(d["_id"])
        d["generated_at"] = d["generated_at"].strftime("%d %b %Y, %H:%M UTC")
        out.append(d)
    return out


def aggregate_quarterly(year: int, quarter: int) -> dict:
    """
    Aggregate totals across all monthly reports for a given quarter.
    Quarter: 1 = Jan-Mar, 2 = Apr-Jun, 3 = Jul-Sep, 4 = Oct-Dec
    """
    db     = get_db()
    months = list(range((quarter - 1) * 3 + 1, (quarter - 1) * 3 + 4))
    pipeline = [
        {"$match": {"period_year": year, "period_month": {"$in": months}}},
        {"$group": {
            "_id":            None,
            "total_codified": {"$sum": "$totals.total_codified"},
            "fwd_dca":        {"$sum": "$totals.fwd_dca"},
            "nsn_allotted":   {"$sum": "$totals.nsn_allotted"},
            "returned":       {"$sum": "$totals.returned"},
            "months_covered": {"$addToSet": "$period_month"},
            "report_count":   {"$sum": 1},
        }},
    ]
    result = list(db.reports.aggregate(pipeline))
    if not result:
        return {"total_codified": 0, "fwd_dca": 0,
                "nsn_allotted": 0, "returned": 0,
                "months_covered": [], "report_count": 0}
    r = result[0]
    r.pop("_id", None)
    return r


def aggregate_yearly(year: int) -> dict:
    """Aggregate totals across all 12 months for a given year."""
    db = get_db()
    pipeline = [
        {"$match": {"period_year": year}},
        {"$group": {
            "_id":            None,
            "total_codified": {"$sum": "$totals.total_codified"},
            "fwd_dca":        {"$sum": "$totals.fwd_dca"},
            "nsn_allotted":   {"$sum": "$totals.nsn_allotted"},
            "returned":       {"$sum": "$totals.returned"},
            "months_covered": {"$addToSet": "$period_month"},
            "report_count":   {"$sum": 1},
        }},
    ]
    result = list(db.reports.aggregate(pipeline))
    if not result:
        return {"total_codified": 0, "fwd_dca": 0,
                "nsn_allotted": 0, "returned": 0,
                "months_covered": [], "report_count": 0}
    r = result[0]
    r.pop("_id", None)
    return r


def aggregate_dpsu(year: int = None) -> list:
    """
    DPSU-wise aggregated totals across all reports (optionally filtered by year).
    Useful for building yearly DPSU comparison charts.
    """
    db    = get_db()
    match = {"period_year": year} if year else {}
    pipeline = [
        {"$match": match},
        {"$unwind": "$rows"},
        {"$group": {
            "_id":            "$rows.dpsu",
            "total_codified": {"$sum": "$rows.total_codified"},
            "fwd_dca":        {"$sum": "$rows.fwd_dca"},
            "nsn_allotted":   {"$sum": "$rows.nsn_allotted"},
            "returned":       {"$sum": "$rows.returned"},
        }},
        {"$sort": {"total_codified": -1}},
    ]
    result = list(db.reports.aggregate(pipeline))
    for r in result:
        r["dpsu"] = r.pop("_id")
    return result