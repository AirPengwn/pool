#!/usr/bin/env python3
"""pool_volume.py - volume of the pool from measured geometry.

Per John's sketch (2026-07-17): the pool is 3 ft deep all around the
perimeter -- the shallow shelf AND every deep-end wall read 3 ft. The deep
end is a hopper: an inverted truncated pyramid scooped out of the deep half,
dropping from the 3 ft rim to a 10x14 flat at 8.5 ft.

    volume = (full footprint filled to 3 ft)      # top slab
           + (hopper below 3 ft)                  # deep scoop
           - stairs displacement
"""

GAL_PER_CUFT = 7.48052

LENGTH           = 40.0
WIDTH            = 20.0
PERIM_DEPTH      = 3.0     # depth at the walls / shallow shelf (measured)
DEEP_DEPTH       = 8.5
SHALLOW_FLAT_LEN = 12.0    # shallow shelf, 0-12 ft
DEEP_FLAT_LEN    = 10.0    # deep flat, along length (x)
DEEP_FLAT_WIDTH  = 14.0    # deep flat, across width (y)
STAIRS_GAL       = 220.0   # semicircle steps ~5 ft wide

BASE_GAL = 34400.0
BASE_PPM = 3.5

# top slab: whole 40x20 footprint filled to the 3 ft perimeter depth
slab = LENGTH * WIDTH * PERIM_DEPTH

# hopper: frustum from the 3 ft rim (opens over the whole deep half) down to
# the 10x14 flat.  Prismatoid formula V = h/6 * (B + 4M + T).
rim_len = LENGTH - SHALLOW_FLAT_LEN            # 28 ft (x = 12..40)
rim_wid = WIDTH                                 # 20 ft
h  = DEEP_DEPTH - PERIM_DEPTH                    # 5.5 ft
B  = DEEP_FLAT_LEN * DEEP_FLAT_WIDTH             # bottom  10x14 = 140
T  = rim_len * rim_wid                            # top     28x20 = 560
M  = ((DEEP_FLAT_LEN + rim_len)/2) * ((DEEP_FLAT_WIDTH + rim_wid)/2)  # mid 323
hopper = (h/6.0) * (B + 4*M + T)

cuft = slab + hopper
gal  = cuft * GAL_PER_CUFT - STAIRS_GAL
ppm  = BASE_PPM * BASE_GAL / gal

print(f"top slab   40x20 @ {PERIM_DEPTH:.0f} ft              {slab*GAL_PER_CUFT:7.0f} gal")
print(f"deep hopper rim {rim_len:.0f}x{rim_wid:.0f} -> flat "
      f"{DEEP_FLAT_LEN:.0f}x{DEEP_FLAT_WIDTH:.0f} ({h:.1f} ft)  {hopper*GAL_PER_CUFT:7.0f} gal")
print(f"stairs                                {-STAIRS_GAL:7.0f} gal")
print("-" * 46)
print(f"TOTAL                                {gal:7.0f} gal")
print(f"\navg depth {cuft/(LENGTH*WIDTH):.2f} ft    dose {ppm:.2f} ppm/gal"
      f"    (baseline {BASE_GAL:.0f} @ {BASE_PPM})")
