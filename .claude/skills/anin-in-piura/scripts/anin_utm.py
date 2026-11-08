"""Conversión y validación de coordenadas para el Proyecto IN Piura.

UTM WGS84 Zona 17S (EPSG:32717) <-> lat/lon WGS84 (EPSG:4326), con pyproj.

Uso en línea de comandos:
    python anin_utm.py latlon -5.2 -79.8
    python anin_utm.py utm 650000 9425000

Uso como módulo:
    from anin_utm import latlon_a_utm, utm_a_latlon, validar_utm
"""
import sys

from pyproj import Transformer

ESTE_MIN, ESTE_MAX = 450_000, 750_000
NORTE_MIN, NORTE_MAX = 9_300_000, 9_600_000

_A_UTM = Transformer.from_crs("EPSG:4326", "EPSG:32717", always_xy=True)
_A_LL = Transformer.from_crs("EPSG:32717", "EPSG:4326", always_xy=True)


def validar_utm(este, norte):
    """Devuelve (ok, mensaje). El mensaje orienta sobre la causa probable."""
    if ESTE_MIN <= este <= ESTE_MAX and NORTE_MIN <= norte <= NORTE_MAX:
        return True, "OK"
    pistas = []
    if ESTE_MIN <= norte <= ESTE_MAX and NORTE_MIN <= este <= NORTE_MAX:
        pistas.append("ESTE y NORTE parecen invertidos")
    if not ESTE_MIN <= este <= ESTE_MAX:
        pistas.append(f"ESTE {este:,.2f} fuera de {ESTE_MIN:,}-{ESTE_MAX:,}")
    if not NORTE_MIN <= norte <= NORTE_MAX:
        pistas.append(f"NORTE {norte:,.2f} fuera de {NORTE_MIN:,}-{NORTE_MAX:,}")
    return False, "; ".join(pistas)


def latlon_a_utm(lat, lon):
    """lat/lon (grados decimales) -> (este, norte, ok, mensaje) en UTM 17S."""
    if abs(lat) > 90:
        raise ValueError(f"lat={lat} fuera de rango: ¿se invirtieron lat y lon?")
    if -82 < lat < -68 and -19 < lon < 0:
        print(f"AVISO: lat={lat}, lon={lon} parecen invertidos (en Perú lon ≈ -81 a -69, lat ≈ -18 a 0).",
              file=sys.stderr)
    if lat > 0:
        # Piura está al sur del Ecuador; una latitud positiva suele ser un signo perdido.
        print(f"AVISO: lat={lat} es positiva; Piura tiene latitud negativa.", file=sys.stderr)
    este, norte = _A_UTM.transform(lon, lat)
    ok, msg = validar_utm(este, norte)
    return este, norte, ok, msg


def utm_a_latlon(este, norte):
    """UTM 17S -> (lat, lon, ok, mensaje)."""
    ok, msg = validar_utm(este, norte)
    lon, lat = _A_LL.transform(este, norte)
    return lat, lon, ok, msg


def _main(argv):
    if len(argv) != 4 or argv[1] not in ("latlon", "utm"):
        print(__doc__)
        return 2
    a, b = float(argv[2]), float(argv[3])
    if argv[1] == "latlon":
        e, n, ok, msg = latlon_a_utm(a, b)
        print(f"ESTE={e:,.2f}  NORTE={n:,.2f}  (UTM 17S WGS84)  [{msg}]")
    else:
        lat, lon, ok, msg = utm_a_latlon(a, b)
        print(f"LAT={lat:.6f}  LON={lon:.6f}  [{msg}]")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(_main(sys.argv))
