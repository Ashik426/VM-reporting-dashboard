"""Builds dashboard/geo.js from geoBoundaries BGD ADM2 (zila) / ADM3 (upazila) simplified GeoJSON.
All 64 zilas and all upazilas are included so any project area can be mapped, and the
upazila->zila gazetteer is used to auto-correct place names in the data.
Usage: python build_geo.py <ADM2.geojson> <ADM3.geojson> <out geo.js>"""
import json, sys
from shapely.geometry import shape, mapping

def rnd(geom, nd):
    g = mapping(geom)
    def r(c):
        return [round(c[0], nd), round(c[1], nd)] if isinstance(c[0], (int, float)) else [r(x) for x in c]
    return {"type": g["type"], "coordinates": r(g["coordinates"])}

adm2 = json.load(open(sys.argv[1], encoding="utf-8"))
adm3 = json.load(open(sys.argv[2], encoding="utf-8"))
dshapes = [(f["properties"]["shapeName"], shape(f["geometry"])) for f in adm2["features"]]

districts = [{"type": "Feature", "properties": {"name": n},
              "geometry": rnd(g.simplify(0.004, preserve_topology=True), 3)} for n, g in dshapes]
upazilas = []
for f in adm3["features"]:
    g = shape(f["geometry"])
    p = g.representative_point()
    # zila = the district polygon containing the upazila (largest overlap as fallback)
    d = next((n for n, dg in dshapes if dg.contains(p)), None) or max(dshapes, key=lambda x: x[1].intersection(g).area)[0]
    upazilas.append({"type": "Feature", "properties": {"name": f["properties"]["shapeName"], "district": d},
                     "geometry": rnd(g.simplify(0.0025, preserve_topology=True), 3)})

out = {"districts": {"type": "FeatureCollection", "features": districts},
       "upazilas": {"type": "FeatureCollection", "features": upazilas}}
open(sys.argv[3], "w", encoding="utf-8").write(
    "// Boundaries: geoBoundaries (CC BY 4.0) BGD ADM2/ADM3, simplified. Built by tools/build_geo.py\nwindow.GEO=" +
    json.dumps(out, separators=(",", ":"), ensure_ascii=False) + ";\n")
print(len(districts), len(upazilas))
