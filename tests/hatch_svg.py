# -*- coding: utf-8 -*-
"""
Created on Wed Jun  4 12:09:47 2025

@author: user
"""

import math

# Pattern parameters
pattern_width = 20         # width of the pattern tile (mm)
pattern_height = 20        # height of the pattern tile (mm)
spacing = 3                # spacing between lines (mm)
angle_deg = 45        # ANSI37: 37.5 degrees
angle_rad = math.radians(angle_deg)

# Function to convert mm to px (Inkscape uses 96 px/inch)
def mm_to_px(mm):
    return mm * 96 / 25.4

width_px = mm_to_px(pattern_width)
height_px = mm_to_px(pattern_height)
spacing_px = mm_to_px(spacing)

# Calculate line length so it covers the pattern box at an angle
line_length = math.hypot(width_px, height_px) * 1.2

# SVG preamble
svg = [
    f'<svg xmlns="http://www.w3.org/2000/svg" width="{pattern_width}mm" height="{pattern_height}mm" viewBox="0 0 {width_px} {height_px}">',
    f'  <g id="ansi37_hatch">'
]

# Draw parallel lines at 37.5°
# We want the lines to fill the bounding box
# We'll move the starting point by 'spacing' along a vector perpendicular to the angle

# Perpendicular direction vector (normalized)
dx = spacing_px * math.sin(angle_rad)
dy = -spacing_px * math.cos(angle_rad)

# Find out how many lines we need to cover the box
num_lines = int((width_px + height_px) / spacing_px) + 2

for i in range(-num_lines, num_lines):
    # Compute line start (x0, y0) at top or left edge, then extend across the tile
    # Offset from (0,0) along perpendicular direction
    offset_x = i * dx
    offset_y = i * dy

    # Line center point
    x0 = offset_x
    y0 = offset_y

    # Calculate line endpoints
    x1 = x0 - math.cos(angle_rad) * line_length/2
    y1 = y0 - math.sin(angle_rad) * line_length/2
    x2 = x0 + math.cos(angle_rad) * line_length/2
    y2 = y0 + math.sin(angle_rad) * line_length/2

    svg.append(f'    <line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" stroke="black" stroke-width="1"/>')

svg.append('  </g>')
svg.append('</svg>')

# Write to file
with open("ansi37_hatch.svg", "w") as f:
    f.write('\n'.join(svg))

print("Generated ansi37_hatch.svg! Import it into Inkscape and use as a pattern.")
