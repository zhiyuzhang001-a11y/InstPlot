# Real-format regression samples

These six user-provided, non-personal measurement files exercise formats that
synthetic fixtures do not fully represent:

- GB2312/GBK Chinese headers;
- CRLF and trailing tab delimiters;
- tab-separated headers containing spaces;
- a CSV containing both text timestamps and numeric columns;
- a Quantum Design VSM `[Header]`/`[Data]` file.

They are retained only for parser regression tests. File names and measurement
values are preserved so future changes can reproduce the original import
behavior. Do not replace them with private or identifying laboratory data.
