# odc-batch-excel

[![CI](https://github.com/jyunji-watanabe/odc-batch-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/jyunji-watanabe/odc-batch-excel/actions/workflows/ci.yml)

An [OutSystems Developer Cloud](https://www.outsystems.com/developer-cloud/) External Logic library (C#, .NET 10)
that will provide batch processing of Excel files for ODC apps. For now, it contains a single sample action, `Greet`.

## Quick start

```sh
dotnet test -c Release   # build, analyzers, tests
scripts/package.sh       # creates artifacts/OdcBatchExcel.zip
```

Upload `artifacts/OdcBatchExcel.zip` in the ODC Portal (**External Logic**) and publish it.

See [AGENTS.md](AGENTS.md) for all commands, conventions and the ODC deployment flow.

## License

[BSD 3-Clause](LICENSE)
