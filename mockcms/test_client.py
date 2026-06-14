#!/usr/bin/env python3
"""
Exercise the mock CMS server the same way the VBA client (modCms.bas) does:
builds byte-compatible SOAP GET/SET envelopes, fires several concurrently,
and prints the parsed curves.

Usage:
    python test_client.py [base_url]
    (default base_url: http://localhost:8080/CreditMarkingService/)
"""

import asyncio
import sys
from xml.sax.saxutils import escape as xesc
import xml.etree.ElementTree as ET

import aiohttp

BASE = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8080/CreditMarkingService/"
NS = "http://www.lehman.com/fta/CreditMarkingService"

USER_META = (
    "<user-metadata><user-details>"
    "<user-id>tester</user-id><user-domain>DEV</user-domain>"
    "<application-id>CMS - Excel Addin</application-id>"
    "<application-version>5.0.1</application-version>"
    "<id>tester</id><host-id>DEVBOX</host-id>"
    "</user-details></user-metadata>"
)

TENORS_17 = ["0M", "3M", "6M", "9M", "1Y", "2Y", "3Y", "4Y",
             "5Y", "6Y", "7Y", "8Y", "9Y", "10Y", "15Y", "20Y", "30Y"]
TENORS_12 = ["0M", "3M", "6M", "9M", "1Y", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "10Y"]


def soap_envelope(operation, rosetta):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="no"?>'
        '<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/">'
        '<SOAP-ENV:Body>'
        f'<SOAPSDK4:{operation} xmlns:SOAPSDK4="{NS}">'
        f"<string>{xesc(USER_META)}</string>"
        f"<string0>{xesc(rosetta)}</string0>"
        f"</SOAPSDK4:{operation}>"
        "</SOAP-ENV:Body></SOAP-ENV:Envelope>"
    )


def get_rosetta(label, ccy="USD", tag="LIVE", date="20260614"):
    return (
        '<Rosetta version="5.0.15"><market><marketData>'
        f"<action>GET</action><date>{date}</date>"
        "<marketSet>"
        "<location>NYC</location>"
        f"<currency>{ccy}</currency><dataSource>{tag}</dataSource>"
        "<version>CLOSE</version><type>creditSpread</type>"
        f"<label>{label}</label>"
        "<genericKeys>"
        f'<nameValuePair name="tag">{tag}</nameValuePair>'
        '<nameValuePair name="returnCreditCurveType">SPREADS</nameValuePair>'
        "</genericKeys></marketSet>"
        "</marketData></market></Rosetta>"
    )


def set_rosetta(label, ticker, ccy, debtclass, product, quotes12, date="20260614"):
    points = []
    for tenor, val in zip(TENORS_12, quotes12):
        if val is None:
            continue
        pm, per = tenor[:-1], tenor[-1]
        points.append(
            f"<point><label>{label}</label><creditSpread>"
            f"<issuerTicker>{ticker}</issuerTicker><debtClass>{debtclass}</debtClass>"
            f"<currency>{ccy}</currency><creditCurveType>{product}</creditCurveType>"
            f"<periodMultiplier>{pm}</periodMultiplier><period>{per}</period>"
            "<contractualSpread>-777</contractualSpread></creditSpread>"
            f"<value>{val}</value></point>"
        )
    return (
        '<Rosetta version="5.0.15"><market><marketData>'
        f"<action>SET</action><date>{date}</date>"
        "<marketSet><location>NYC</location><logicalTime>LIVE</logicalTime>"
        f"{''.join(points)}"
        "</marketSet></marketData></market></Rosetta>"
    )


def _local(tag):
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def parse_response(soap_text, is_set):
    env = ET.fromstring(soap_text)
    # The client takes the Body's text content -> the escaped Rosetta.
    body = next(e for e in env.iter() if _local(e.tag) == "Body")
    rosetta_text = "".join(body.itertext()).strip()
    rose = ET.fromstring(rosetta_text)

    out = {}
    for ms in (e for e in rose.iter() if _local(e.tag) == "marketSet"):
        label = next((c.text for c in ms if _local(c.tag) == "label"), "?")
        quotes = {}
        for pt in (e for e in ms if _local(e.tag) == "point"):
            val = next((c.text for c in pt if _local(c.tag) == "value"), None)
            spread = next((c for c in pt if _local(c.tag) == "creditSpread"), None)
            if spread is None or val is None:
                continue
            pm = next((c.text for c in spread if _local(c.tag) == "periodMultiplier"), "")
            per = next((c.text for c in spread if _local(c.tag) == "period"), "")
            tenor = f"{pm}{per}"
            # GET: x10000 -> bps ; SET: verbatim
            quotes[tenor] = float(val) if is_set else float(val) * 10000.0
        out[label] = quotes
    return out


async def soap_call(session, operation, rosetta):
    headers = {"Content-Type": "text/xml; charset=utf-8", "SOAPAction": '""'}
    async with session.post(BASE, data=soap_envelope(operation, rosetta),
                            headers=headers) as resp:
        return await resp.text()


async def main():
    labels = [
        "AEP.USD.SENIOR_NORE_14.SNAC100",
        "F.USD.SENIOR_NORE_14.SNAC100",
        "T.USD.SENIOR_NORE_14.SNAC100",
        "IBM.USD.SENIOR_NORE_14.SNAC100",
    ]
    async with aiohttp.ClientSession() as s:
        print("=== concurrent first GETs (random curves generated) ===")
        gets = [soap_call(s, "getMarketData", get_rosetta(l)) for l in labels]
        for label, text in zip(labels, await asyncio.gather(*gets)):
            q = parse_response(text, is_set=False)[label]
            print(f"{label:38s}  5Y={q['5Y']:7.2f}  10Y={q['10Y']:7.2f}  "
                  f"1Y={q['1Y']:7.2f}  (bps)")

        print("\n=== SET a new curve on AEP (mark 5Y=200) ===")
        marked = [12, 25, 40, 55, 80, 130, 165, 188, 200, 208, 214, 226]
        text = await soap_call(s, "setMarketData",
                               set_rosetta(labels[0], "AEP", "USD",
                                           "SENIOR_NORE_14", "SNAC100", marked))
        echo = parse_response(text, is_set=True)[labels[0]]
        print(f"SET echo                                5Y={echo['5Y']:7.2f}  "
              f"10Y={echo['10Y']:7.2f}  (bps, verbatim)")

        print("\n=== GET AEP again - reflects the mark ===")
        text = await soap_call(s, "getMarketData", get_rosetta(labels[0]))
        q = parse_response(text, is_set=False)[labels[0]]
        print(f"{labels[0]:38s}  5Y={q['5Y']:7.2f}  10Y={q['10Y']:7.2f}  (bps)")
        assert abs(q["5Y"] - 200) < 1e-6, "marked 5Y should round-trip"
        print("\nOK - SET round-tripped through a subsequent GET.")


if __name__ == "__main__":
    asyncio.run(main())
