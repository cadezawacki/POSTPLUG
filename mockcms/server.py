#!/usr/bin/env python3
"""
Mock CMS Credit Marking Service - a local stand-in for the production
CreditMarkingService used by the ExcelBridge CMS engine (modCms.bas).

It speaks the exact wire contract the VBA client expects:

  * SOAP transport (default) - rpc/encoded SOAP 1.1 envelope POSTed to any
    path (e.g. /CreditMarkingService/). The body carries two string params:
        <string>   XML-escaped user-metadata
        <string0>  XML-escaped Rosetta GET/SET request
  * REST transport (the modernized facade) - JSON POSTed to
        /api/v1/getMarketData   /api/v1/setMarketData
    with the same Rosetta inside requestHeader/request.

GET vs SET is decided by the Rosetta <action> element, exactly as the real
service routes it.

Behaviour designed for development:
  * The first GET for a curve we have never seen returns a freshly generated,
    random-but-realistically-shaped CDS curve (positive, upward sloping, 5Y a
    rounded integer in bps). It is then cached, so repeat GETs are stable.
  * A SET stores the marked levels and echoes them back verbatim, just like
    production; the next GET reflects them. So you can mark/re-mark all day
    without touching real curves.
  * Everything is async (aiohttp) - many GET/SET requests are served
    concurrently on one event loop, with optional simulated latency so the
    parallelism behaves like the real network.

Units (must match modCms.bas):
  GET response : creditSpread value x 10000 -> bps ; creditUpfront x 100 -> %
  SET request  : values are verbatim (already in display units)
  => GET wire value = display / 10000 (spread) or / 100 (upfront)
     SET wire value = display verbatim

Run:  python server.py            (SOAP + REST on http://0.0.0.0:8080)
      python server.py --help     for options
"""

import argparse
import asyncio
import json
import logging
import random
from datetime import datetime
from xml.sax.saxutils import escape as xml_escape
import xml.etree.ElementTree as ET

from aiohttp import web

# --------------------------------------------------------------------------- #
#  Wire constants - mirror modCms.bas exactly
# --------------------------------------------------------------------------- #

ROSETTA_VERSION = "5.0.15"
CMS_NAMESPACE = "http://www.lehman.com/fta/CreditMarkingService"

# 17-standard GET grid, in order.
TENORS_17 = ["0M", "3M", "6M", "9M", "1Y", "2Y", "3Y", "4Y",
             "5Y", "6Y", "7Y", "8Y", "9Y", "10Y", "15Y", "20Y", "30Y"]

# 12-standard SET grid (the only tenors the client ever marks).
TENORS_12 = ["0M", "3M", "6M", "9M", "1Y", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "10Y"]

# Typical upward-sloping CDS term structure as a ratio to the 5Y point.
# (Front end well below 5Y, long end modestly above it - the normal shape.)
SHAPE_RATIO = {
    "0M": 0.12, "3M": 0.18, "6M": 0.26, "9M": 0.33, "1Y": 0.42,
    "2Y": 0.58, "3Y": 0.72, "4Y": 0.87, "5Y": 1.00, "6Y": 1.05,
    "7Y": 1.09, "8Y": 1.12, "9Y": 1.15, "10Y": 1.18, "15Y": 1.22,
    "20Y": 1.25, "30Y": 1.27,
}

log = logging.getLogger("mockcms")


# --------------------------------------------------------------------------- #
#  Curve generation + in-memory store
# --------------------------------------------------------------------------- #

class CurveStore:
    """Holds the simulated market. Keyed by (label, tag, date) so LIVE and
    historical (NYOISCLOSE) snapshots are independent, just like production."""

    def __init__(self):
        self._curves = {}
        self._lock = asyncio.Lock()

    @staticmethod
    def _key(label, tag, date):
        return (label.upper(), (tag or "LIVE").upper(), date or "")

    async def get_or_generate(self, label, ccy, tag, date, field):
        """Return the stored curve, generating + caching one on first sight."""
        async with self._lock:
            k = self._key(label, tag, date)
            cv = self._curves.get(k)
            if cv is None:
                cv = generate_curve(label, ccy, field)
                self._curves[k] = cv
                log.info("generated new curve  %s  (5Y=%s %s)",
                         label, cv["quotes"].get("5Y"), cv["field"])
            return cv

    async def apply_set(self, label, ccy, tag, date, field, set_quotes, recovery, user):
        """Merge marked levels onto a curve (generating a full one first if we
        have never seen it) and return the stored curve."""
        async with self._lock:
            k = self._key(label, tag, date)
            cv = self._curves.get(k)
            if cv is None:
                cv = generate_curve(label, ccy, field)
                self._curves[k] = cv
            cv["field"] = field
            for tenor, val in set_quotes.items():
                cv["quotes"][tenor] = val
            if recovery is not None and recovery > 0:
                cv["recovery"] = recovery
            cv["owner"] = user or cv["owner"]
            cv["lastMarkedBy"] = user or cv["lastMarkedBy"]
            cv["marked"] = _now_stamp()
            cv["markedType"] = "QUOTED" if field == "creditUpfront" else "SPREAD"
            log.info("marked curve         %s  (%d tenors, by %s)",
                     label, len(set_quotes), cv["lastMarkedBy"])
            return cv


def _now_stamp():
    # CMS hands back "2008-12-12 17:04:41.137"; modCms.ConvertToDate parses it.
    now = datetime.now()
    return now.strftime("%Y-%m-%d %H:%M:%S.") + f"{now.microsecond // 1000:03d}"


def generate_curve(label, ccy, field):
    """A random but realistically shaped CDS curve.

    Picks a credit regime, draws a 5Y level (a rounded integer in bps), then
    fans the other tenors out around it with the standard upward slope plus a
    little per-tenor jitter. All values are strictly positive.
    """
    regime = random.choices(["IG", "XO", "HY"], weights=[0.50, 0.35, 0.15])[0]
    if regime == "IG":
        five_yr = random.randint(25, 120)
    elif regime == "XO":
        five_yr = random.randint(120, 300)
    else:
        five_yr = random.randint(300, 650)

    quotes = {}
    for tenor in TENORS_17:
        if tenor == "5Y":
            quotes[tenor] = float(five_yr)            # rounded integer
            continue
        jitter = random.uniform(0.97, 1.03)
        val = five_yr * SHAPE_RATIO[tenor] * jitter
        quotes[tenor] = round(max(val, 0.01), 2)      # positive, 2dp

    if field == "creditUpfront":
        # Convert the bps-shaped curve into plausible upfront % points so the
        # UPFRONT convention path returns something sensible too.
        quotes = {t: round(v / 100.0, 4) for t, v in quotes.items()}
        quotes["5Y"] = round(quotes["5Y"], 2)

    return {
        "field": field,
        "quotes": quotes,
        "recovery": random.choice([0.40, 0.40, 0.40, 0.25]),
        "owner": "mockuser",
        "lastMarkedBy": "mockuser",
        "marked": _now_stamp(),
        "isParent": "true",
        "markedType": "QUOTED" if field == "creditUpfront" else "SPREAD",
    }


# --------------------------------------------------------------------------- #
#  Small XML helpers (namespace-agnostic)
# --------------------------------------------------------------------------- #

def _local(tag):
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _find_local(root, name):
    if _local(root.tag) == name:
        return root
    for e in root.iter():
        if _local(e.tag) == name:
            return e
    return None


def _find_all_local(root, name):
    return [e for e in root.iter() if _local(e.tag) == name]


def _text_of(root, name, default=""):
    e = _find_local(root, name)
    return (e.text or default) if e is not None else default


def _split_tenor(tenor):
    """'10Y' -> ('10', 'Y');  '0M' -> ('0', 'M')."""
    return tenor[:-1], tenor[-1]


def _fmt_num(x):
    """Locale-free number for the wire; trims trailing zeros."""
    s = f"{x:.10f}".rstrip("0").rstrip(".")
    return s if s else "0"


# --------------------------------------------------------------------------- #
#  Request parsing
# --------------------------------------------------------------------------- #

def extract_user_id(soap_string_meta):
    if not soap_string_meta:
        return None
    try:
        meta = ET.fromstring(soap_string_meta)
        uid = _find_local(meta, "user-id")
        return uid.text if uid is not None else None
    except ET.ParseError:
        return None


def parse_rosetta(rosetta_xml):
    """Parse a GET or SET Rosetta request into a normalized dict:

      GET -> {"action": "GET", "date": ..., "curves": [
                  {"label","ccy","tag","field"}, ... ]}
      SET -> {"action": "SET", "date": ..., "curves": [
                  {"label","ccy","field","quotes":{tenor:val},"recovery"} ]}
    """
    root = ET.fromstring(rosetta_xml)
    market_data = _find_local(root, "marketData")
    if market_data is None:
        raise ValueError("no <marketData> in Rosetta")

    action = (_text_of(market_data, "action") or "").upper()
    date = _text_of(market_data, "date")
    market_sets = _find_all_local(market_data, "marketSet")

    if action == "GET":
        curves = []
        for ms in market_sets:
            label = _text_of(ms, "label")
            ccy = _text_of(ms, "currency")
            tag = ""
            ret_type = "SPREADS"
            gk = _find_local(ms, "genericKeys")
            if gk is not None:
                for nvp in _find_all_local(gk, "nameValuePair"):
                    name = (nvp.get("name") or "").lower()
                    if name == "tag":
                        tag = nvp.text or ""
                    elif name == "returncreditcurvetype":
                        ret_type = (nvp.text or "SPREADS")
            if not tag:
                tag = _text_of(ms, "dataSource") or "LIVE"
            field = "creditUpfront" if "UPFRONT" in ret_type.upper() else "creditSpread"
            if label:
                curves.append({"label": label, "ccy": ccy, "tag": tag, "field": field})
        return {"action": "GET", "date": date, "curves": curves}

    if action == "SET":
        # One curve per SET request: many <point>s in a single <marketSet>.
        curves = []
        for ms in market_sets:
            label = ""
            ccy = ""
            field = "creditSpread"
            quotes = {}
            recovery = None
            for pt in _find_all_local(ms, "point"):
                val_el = _find_local(pt, "value")
                val = float(val_el.text) if val_el is not None and val_el.text else None
                lbl_el = _first_child_local(pt, "label")
                if lbl_el is not None and lbl_el.text:
                    label = lbl_el.text
                spread = _first_child_local(pt, "creditSpread")
                upfront = _first_child_local(pt, "creditUpfront")
                recov = _first_child_local(pt, "issuerRecovery")
                quote_el = spread if spread is not None else upfront
                if quote_el is not None and val is not None:
                    field = _local(quote_el.tag)
                    if not ccy:
                        ccy = _text_of(quote_el, "currency")
                    pm = _text_of(quote_el, "periodMultiplier")
                    per = _text_of(quote_el, "period")
                    tenor = f"{pm}{per}"
                    if tenor in TENORS_17:
                        quotes[tenor] = val
                elif recov is not None and val is not None:
                    recovery = val
                    if not ccy:
                        ccy = _text_of(recov, "currency")
            if label:
                curves.append({"label": label, "ccy": ccy, "field": field,
                               "quotes": quotes, "recovery": recovery})
        return {"action": "SET", "date": date, "curves": curves}

    raise ValueError(f"unsupported action '{action}'")


def _first_child_local(parent, name):
    for child in list(parent):
        if _local(child.tag) == name:
            return child
    return None


# --------------------------------------------------------------------------- #
#  Response building
# --------------------------------------------------------------------------- #

def _wire_value(display, field, is_set_echo):
    if is_set_echo:
        return _fmt_num(display)               # echoed verbatim
    if field == "creditUpfront":
        return _fmt_num(display / 100.0)        # % -> fraction
    return _fmt_num(display / 10000.0)          # bps -> fraction


def _market_set_xml(label, cv, tenors, is_set_echo):
    field = cv["field"]
    points = []
    for tenor in tenors:
        if tenor not in cv["quotes"]:
            continue
        pm, per = _split_tenor(tenor)
        wire = _wire_value(cv["quotes"][tenor], field, is_set_echo)
        points.append(
            f"<point><label>{xml_escape(label)}</label>"
            f"<{field}>"
            f"<periodMultiplier>{pm}</periodMultiplier><period>{per}</period>"
            f"</{field}>"
            f"<value>{wire}</value></point>"
        )
    # Recovery point (never scaled in either direction).
    if cv.get("recovery"):
        points.append(
            f"<point><label>{xml_escape(label)}</label>"
            f"<issuerRecovery></issuerRecovery>"
            f"<value>{_fmt_num(cv['recovery'])}</value></point>"
        )

    generic = (
        "<genericKeys>"
        f"<nameValuePair name=\"owner\">{xml_escape(cv['owner'])}</nameValuePair>"
        f"<nameValuePair name=\"marked\">{xml_escape(cv['marked'])}</nameValuePair>"
        f"<nameValuePair name=\"lastMarkedBy\">{xml_escape(cv['lastMarkedBy'])}</nameValuePair>"
        f"<nameValuePair name=\"isParent\">{xml_escape(cv['isParent'])}</nameValuePair>"
        f"<nameValuePair name=\"latestCurveMarkedType\">{xml_escape(cv['markedType'])}</nameValuePair>"
        "</genericKeys>"
    )
    return (
        "<marketSet>"
        f"<label>{xml_escape(label)}</label>"
        f"{generic}"
        f"{''.join(points)}"
        "</marketSet>"
    )


def build_rosetta_response(parsed, store_results, is_set):
    """Assemble the full Rosetta response document the client will parse."""
    tenors = TENORS_12 if is_set else TENORS_17
    market_sets = "".join(
        _market_set_xml(label, cv, tenors, is_set_echo=is_set)
        for (label, cv) in store_results
    )
    system_messages = (
        "<systemMessages><statusMessage>"
        "<statusCode>OK</statusCode>"
        "<statusDescription>Success</statusDescription>"
        "</statusMessage></systemMessages>"
    )
    action = "SET" if is_set else "GET"
    return (
        f'<Rosetta version="{ROSETTA_VERSION}"><market><marketData>'
        f"<action>{action}</action>"
        f"{system_messages}"
        f"{market_sets}"
        "</marketData></market></Rosetta>"
    )


def build_error_rosetta(message):
    return (
        f'<Rosetta version="{ROSETTA_VERSION}"><market><marketData>'
        "<systemMessages><statusMessage>"
        "<statusCode>ERROR</statusCode>"
        f"<statusDescription>{xml_escape(message)}</statusDescription>"
        "</statusMessage></systemMessages>"
        "<systemErrors><errorMessage>"
        f"<errorMessageContent>{xml_escape(message)}</errorMessageContent>"
        "</errorMessage></systemErrors>"
        "</marketData></market></Rosetta>"
    )


def soap_envelope(operation, rosetta_xml):
    """Wrap a Rosetta payload in an rpc/encoded SOAP response. The client reads
    the Body's text content, so the Rosetta is XML-escaped inside <result>."""
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
        "<soap:Body>"
        f'<ns:{operation}Response xmlns:ns="{CMS_NAMESPACE}">'
        f"<result>{xml_escape(rosetta_xml)}</result>"
        f"</ns:{operation}Response>"
        "</soap:Body></soap:Envelope>"
    )


def soap_fault(message):
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
        "<soap:Body><soap:Fault>"
        "<faultcode>soap:Server</faultcode>"
        f"<faultstring>{xml_escape(message)}</faultstring>"
        "</soap:Fault></soap:Body></soap:Envelope>"
    )


def rest_response(rosetta_xml):
    return json.dumps({
        "responseHeader": {"code": "SUCCESS", "description": "OK"},
        "code": "SUCCESS",
        "response": rosetta_xml,
    })


# --------------------------------------------------------------------------- #
#  Core request handling (transport-independent)
# --------------------------------------------------------------------------- #

async def process_rosetta(app, rosetta_xml, user_id):
    parsed = parse_rosetta(rosetta_xml)
    store = app["store"]
    is_set = parsed["action"] == "SET"

    # Simulate per-request processing/network latency so concurrent requests
    # actually overlap on the event loop.
    lat = app["latency"]
    if lat > 0:
        await asyncio.sleep(random.uniform(lat * 0.5, lat))

    results = []
    if is_set:
        for c in parsed["curves"]:
            cv = await store.apply_set(c["label"], c["ccy"], "LIVE", parsed["date"],
                                       c["field"], c["quotes"], c["recovery"], user_id)
            results.append((c["label"], cv))
    else:
        for c in parsed["curves"]:
            cv = await store.get_or_generate(c["label"], c["ccy"], c["tag"],
                                             parsed["date"], c["field"])
            results.append((c["label"], cv))

    return parsed, build_rosetta_response(parsed, results, is_set)


# --------------------------------------------------------------------------- #
#  HTTP handlers
# --------------------------------------------------------------------------- #

async def handle_soap(request):
    body = await request.text()
    try:
        envelope = ET.fromstring(body)
    except ET.ParseError as exc:
        return web.Response(status=500, content_type="text/xml",
                            text=soap_fault(f"malformed SOAP envelope: {exc}"))

    string0 = _find_local(envelope, "string0")
    if string0 is None or not (string0.text or "").strip():
        return web.Response(status=500, content_type="text/xml",
                            text=soap_fault("no <string0> Rosetta payload in request"))

    meta = _find_local(envelope, "string")
    user_id = extract_user_id(meta.text if meta is not None else None)

    try:
        parsed, rosetta = await process_rosetta(request.app, string0.text, user_id)
    except Exception as exc:                                   # noqa: BLE001
        log.exception("SOAP request failed")
        return web.Response(status=500, content_type="text/xml",
                            text=soap_fault(str(exc)))

    op = "setMarketData" if parsed["action"] == "SET" else "getMarketData"
    log.info("SOAP %-3s  %2d curve(s)  ->  %s", parsed["action"],
             len(parsed["curves"]), op)
    return web.Response(content_type="text/xml", charset="utf-8",
                        text=soap_envelope(op, rosetta))


async def handle_rest(request):
    raw = await request.text()
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as exc:
        return web.json_response(
            {"code": "ERROR", "description": f"bad JSON: {exc}"}, status=400)

    inner = payload.get("request", "")
    user_id = (payload.get("requestHeader") or {}).get("userId")
    start = inner.find("<Rosetta")
    end = inner.rfind("</Rosetta>")
    if start == -1 or end == -1:
        return web.json_response(
            {"code": "ERROR", "description": "no <Rosetta> in request"}, status=400)
    rosetta_req = inner[start:end + len("</Rosetta>")]

    try:
        parsed, rosetta = await process_rosetta(request.app, rosetta_req, user_id)
    except Exception as exc:                                   # noqa: BLE001
        log.exception("REST request failed")
        return web.json_response(
            {"code": "ERROR", "description": str(exc),
             "response": build_error_rosetta(str(exc))}, status=200)

    log.info("REST %-3s  %2d curve(s)", parsed["action"], len(parsed["curves"]))
    return web.Response(content_type="application/json", text=rest_response(rosetta))


async def handle_health(request):
    store = request.app["store"]
    return web.json_response({
        "status": "ok",
        "service": "mock-cms",
        "cached_curves": len(store._curves),
    })


# --------------------------------------------------------------------------- #
#  App wiring
# --------------------------------------------------------------------------- #

def make_app(latency):
    app = web.Application(client_max_size=64 * 1024 * 1024)
    app["store"] = CurveStore()
    app["latency"] = latency
    app.router.add_get("/health", handle_health)
    # REST facade (explicit) - registered before the SOAP catch-all.
    app.router.add_post("/api/v1/getMarketData", handle_rest)
    app.router.add_post("/api/v1/setMarketData", handle_rest)
    # SOAP - any other POST path (e.g. /CreditMarkingService/).
    app.router.add_post("/{tail:.*}", handle_soap)
    return app


def main():
    ap = argparse.ArgumentParser(description="Mock CMS Credit Marking Service")
    ap.add_argument("--host", default="0.0.0.0", help="bind host (default 0.0.0.0)")
    ap.add_argument("--port", type=int, default=8080, help="bind port (default 8080)")
    ap.add_argument("--latency", type=float, default=0.05,
                    help="max simulated per-request latency in seconds "
                         "(default 0.05; 0 disables)")
    ap.add_argument("--seed", type=int, default=None,
                    help="RNG seed for reproducible curves")
    ap.add_argument("-q", "--quiet", action="store_true", help="warnings only")
    args = ap.parse_args()

    if args.seed is not None:
        random.seed(args.seed)

    logging.basicConfig(
        level=logging.WARNING if args.quiet else logging.INFO,
        format="%(asctime)s  %(levelname)-5s  %(message)s",
        datefmt="%H:%M:%S",
    )

    log.info("Mock CMS listening on http://%s:%d", args.host, args.port)
    log.info("  SOAP : POST any path, e.g. http://%s:%d/CreditMarkingService/",
             args.host, args.port)
    log.info("  REST : POST http://%s:%d/api/v1/{getMarketData,setMarketData}",
             args.host, args.port)
    log.info("  simulated latency: up to %.0f ms%s",
             args.latency * 1000, "  (disabled)" if args.latency == 0 else "")

    web.run_app(make_app(args.latency), host=args.host, port=args.port,
                print=None)


if __name__ == "__main__":
    main()
