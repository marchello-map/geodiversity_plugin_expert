Geodiversity Calculator v2.1 — Expert Edition

Advanced geodiversity analysis with hybrid landform modelling (vector + geomorphon).
This expert version extends the standard workflow with a powerful, flexible geomorphology engine, allowing users to combine vector-based landform classification and DEM-derived geomorphon analysis into a unified geodiversity framework.

Vector-Based Landform Variety (Expert Mode)
Supports multi-geometry inputs simultaneously:
Lines (e.g. ridges, faults)
Polygons (landform units)
Points (landform features)
Each dataset:
Uses a user-defined classification field
Is internally encoded into unique categorical values
Merged into a global classification system
✔️ Key innovation:
Cross-layer category merging with offset encoding
Prevents class conflicts between datasets
✔️ Computation:
Counts distinct landform types per grid cell
Uses spatial indexing for performance
Works on any geometry type

DEM-Based Geomorphon Analysis
Uses GRASS r.geomorphon algorithm
Automatically derives landform classes from elevation
Configurable parameters:
Search radius
Flatness threshold
Skip distance
Morphological flags

✔️ Produces:

Landform classes (10 standard geomorphon types)
Grid-level zonal statistics of landform diversity

👉 Result fields:
