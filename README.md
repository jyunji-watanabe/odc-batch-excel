# odc-batch-excel

[![CI](https://github.com/jyunji-watanabe/odc-batch-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/jyunji-watanabe/odc-batch-excel/actions/workflows/ci.yml)

An [OutSystems Developer Cloud](https://www.outsystems.com/developer-cloud/) External Logic library (C#, .NET 10)
for batch processing of Excel files in ODC apps. Its one action, `Execute`, takes an .xlsx workbook and a list of
commands, such as setting values and formulas, inserting or deleting rows and columns, or reading cells and ranges.
It runs them all in one call and returns the updated workbook plus anything the commands read.

See [docs/design.md](docs/design.md) for the contract, the operation catalog and the known limitations.

## Quick start

```sh
dotnet test -c Release   # build, analyzers, tests
scripts/package.sh       # creates artifacts/OdcBatchExcel.zip
```

Upload `artifacts/OdcBatchExcel.zip` in the ODC Portal (**External Logic**) and publish it.

See [AGENTS.md](AGENTS.md) for all commands, conventions and the ODC deployment flow.

## License

[BSD 3-Clause](LICENSE)
