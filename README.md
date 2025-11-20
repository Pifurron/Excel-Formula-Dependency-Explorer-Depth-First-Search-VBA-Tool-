# Excel-Formula-Dependency-Explorer-Depth-First-Search-VBA-Tool-

This VBA module implements a Depth-First Search (DFS) algorithm to analyze formula dependency chains inside Excel workbooks.
When a user selects a cell and runs the tool, it recursively identifies every precedent cell referenced within its formula (including ranges, multi-cell references, cross-sheet references, and chained dependencies), and constructs a hierarchical dependency tree.
