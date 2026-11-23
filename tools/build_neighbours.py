"""Adds neighbouring-country land (India, Myanmar, Bhutan, Nepal, China) around Bangladesh to dashboard/geo.js
as GEO.neighbours, so the printed map shows grey land vs sea like a GIS layout.
Source: Natural Earth 1:50m admin-0 countries (public domain).
Usage: python build_neighbours.py <ne_50m_admin_0_countries.geojson> <geo.js>"""
import json, re, sys
from shapely.geometry import shape, box, mapping

BBOX = box(86.5, 19.3, 94.2, 28.2)
ne = json.load(open(sys.argv[1], encoding="utf-8"))
feats = []
for f in ne["features"]:
    iso = f["properties"].get("ADM0_A3")
    if iso in ("BGD",):
        continue
    g = shape(f["geometry"])
    if not g.intersects(BBOX):
        continue
    g = g.intersection(BBOX).simplify(0.01, preserve_topology=True)
    if g.is_empty:
        continue
    m = mapping(g)
    def r(c):
        return [round(c[0], 3), round(c[1], 3)] if isinstance(c[0], (int, float)) else [r(x) for x in c]
    feats.append({"type": "Feature", "properties": {"name": f["properties"]["NAME"]},
                  "geometry": {"type": m["type"], "coordinates": r(m["coordinates"])}})

src = open(sys.argv[2], encoding="utf-8").read()
src = re.sub(r"\nwindow\.GEO\.neighbours=.*", "", src, flags=re.S).rstrip() + "\n"
src += "window.GEO.neighbours=" + json.dumps({"type": "FeatureCollection", "features": feats}, separators=(",", ":")) + ";\n"
open(sys.argv[2], "w", encoding="utf-8").write(src)
print([f["properties"]["name"] for f in feats])
