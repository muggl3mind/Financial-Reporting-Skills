# Default Useful Lives (IRS MACRS Guidelines)

If `useful_life` is not specified, the system uses these defaults based on asset category:

| Category | Default Useful Life | MACRS Class |
|----------|---------------------|-------------|
| Equipment | 7 years | 7-year property |
| Vehicle | 5 years | 5-year property |
| Furniture | 7 years | 7-year property |
| Computer | 5 years | 5-year property |
| Building | 39 years | Nonresidential real property |
| Software | 3 years | 3-year property |
| Patent | 15 years | 15-year property (intangibles) |

You can override the default by explicitly providing `useful_life` in the asset JSON or `--useful-life` argument.
