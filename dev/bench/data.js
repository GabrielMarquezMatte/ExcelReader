window.BENCHMARK_DATA = {
  "lastUpdate": 1782452802019,
  "repoUrl": "https://github.com/GabrielMarquezMatte/ExcelReader",
  "entries": {
    "Benchmark": [
      {
        "commit": {
          "author": {
            "email": "gabrielandremarquez.matte@gmail.com",
            "name": "Gabriel Matte",
            "username": "GabrielMarquezMatte"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "281b113ec2224d0fa5cd6c2520f15821988a9424",
          "message": "Merge pull request #4 from GabrielMarquezMatte/develop\n\nEnhance benchmark configuration by adding exporters for GitHub and JSON",
          "timestamp": "2026-06-24T00:35:59-03:00",
          "tree_id": "bc496f7620f927e0e278b22e0b4d9197d367619d",
          "url": "https://github.com/GabrielMarquezMatte/ExcelReader/commit/281b113ec2224d0fa5cd6c2520f15821988a9424"
        },
        "date": 1782272267634,
        "tool": "benchmarkdotnet",
        "benches": [
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserSync(Rows: 50000)",
            "value": 39597421.752136745,
            "unit": "ns",
            "range": "± 73692.75777996871"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserAsync(Rows: 50000)",
            "value": 46721223.125,
            "unit": "ns",
            "range": "± 157531.26583651121"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.MiniExcel(Rows: 50000)",
            "value": 338219760.6666667,
            "unit": "ns",
            "range": "± 2009721.5823591088"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.Sylvan(Rows: 50000)",
            "value": 111416861.1111111,
            "unit": "ns",
            "range": "± 610111.4557202734"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.SylvanAsync(Rows: 50000)",
            "value": 111628413,
            "unit": "ns",
            "range": "± 1612203.5865796988"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReader(Rows: 50000)",
            "value": 40209438.800000004,
            "unit": "ns",
            "range": "± 559495.6564649302"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReaderAsync(Rows: 50000)",
            "value": 42067249.3,
            "unit": "ns",
            "range": "± 178531.7350087073"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.MiniExcel(Rows: 50000)",
            "value": 324911620.4444444,
            "unit": "ns",
            "range": "± 1716286.8076147058"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.Sylvan(Rows: 50000)",
            "value": 76472992.02857141,
            "unit": "ns",
            "range": "± 674657.5271946148"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.ExcelReaderWriter(Rows: 50000)",
            "value": 48336470.47474747,
            "unit": "ns",
            "range": "± 140008.3388423644"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.MiniExcel(Rows: 50000)",
            "value": 106505455.8888889,
            "unit": "ns",
            "range": "± 903809.1409505189"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "gabrielandremarquez.matte@gmail.com",
            "name": "Gabriel Matte",
            "username": "GabrielMarquezMatte"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "53897c45268ff2d041449d9cc53ed2d40753bac7",
          "message": "Merge pull request #5 from GabrielMarquezMatte/develop",
          "timestamp": "2026-06-24T02:29:36-03:00",
          "tree_id": "cea983699832650477dd48f14be06796b11bcc7f",
          "url": "https://github.com/GabrielMarquezMatte/ExcelReader/commit/53897c45268ff2d041449d9cc53ed2d40753bac7"
        },
        "date": 1782279076422,
        "tool": "benchmarkdotnet",
        "benches": [
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserSync(Rows: 50000)",
            "value": 37971518.83035714,
            "unit": "ns",
            "range": "± 152812.67502170912"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserAsync(Rows: 50000)",
            "value": 43590147.777777776,
            "unit": "ns",
            "range": "± 255556.11095047588"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.MiniExcel(Rows: 50000)",
            "value": 316811449,
            "unit": "ns",
            "range": "± 1893198.9012877499"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.Sylvan(Rows: 50000)",
            "value": 100144541.33333333,
            "unit": "ns",
            "range": "± 572889.5060225402"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.SylvanAsync(Rows: 50000)",
            "value": 105882601.2,
            "unit": "ns",
            "range": "± 2808831.990338523"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReader(Rows: 50000)",
            "value": 37805109.569230765,
            "unit": "ns",
            "range": "± 204852.6291230216"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReaderAsync(Rows: 50000)",
            "value": 42162785.30555555,
            "unit": "ns",
            "range": "± 124587.67926013189"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.MiniExcel(Rows: 50000)",
            "value": 327795014,
            "unit": "ns",
            "range": "± 1679279.0077513922"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.Sylvan(Rows: 50000)",
            "value": 75042621.85714285,
            "unit": "ns",
            "range": "± 217900.2983812993"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.ExcelReaderWriter(Rows: 50000)",
            "value": 49856747.37373737,
            "unit": "ns",
            "range": "± 102388.37765230697"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.MiniExcel(Rows: 50000)",
            "value": 105656205.25,
            "unit": "ns",
            "range": "± 2284608.687491877"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "gabrielandremarquez.matte@gmail.com",
            "name": "Gabriel Matte",
            "username": "GabrielMarquezMatte"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "803ad790c436ccc6b81f2ad47482df5c76a7c8cb",
          "message": "Merge pull request #6 from GabrielMarquezMatte/develop\n\nDevelop",
          "timestamp": "2026-06-24T12:47:56-03:00",
          "tree_id": "dfdc755969638944d4511eb97dec85c69a74d7e9",
          "url": "https://github.com/GabrielMarquezMatte/ExcelReader/commit/803ad790c436ccc6b81f2ad47482df5c76a7c8cb"
        },
        "date": 1782316175164,
        "tool": "benchmarkdotnet",
        "benches": [
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserSync(Rows: 50000)",
            "value": 38629089.931623936,
            "unit": "ns",
            "range": "± 116383.27329658918"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserAsync(Rows: 50000)",
            "value": 44080616.166666664,
            "unit": "ns",
            "range": "± 1262811.6398333372"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.MiniExcel(Rows: 50000)",
            "value": 312708537.6,
            "unit": "ns",
            "range": "± 1315405.7420717676"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.Sylvan(Rows: 50000)",
            "value": 99018089.125,
            "unit": "ns",
            "range": "± 890214.0430059403"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.SylvanAsync(Rows: 50000)",
            "value": 105279911,
            "unit": "ns",
            "range": "± 713925.5672790173"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReader(Rows: 50000)",
            "value": 37620673.78571428,
            "unit": "ns",
            "range": "± 155043.1565778607"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReaderAsync(Rows: 50000)",
            "value": 41269554.13076923,
            "unit": "ns",
            "range": "± 89575.17894480555"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.MiniExcel(Rows: 50000)",
            "value": 309731245.2,
            "unit": "ns",
            "range": "± 1688351.2780070251"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.Sylvan(Rows: 50000)",
            "value": 71064334.28571428,
            "unit": "ns",
            "range": "± 302456.9111672867"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.ExcelReaderWriter(Rows: 50000)",
            "value": 47885032.16161617,
            "unit": "ns",
            "range": "± 172564.86280827518"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.MiniExcel(Rows: 50000)",
            "value": 110619082.6,
            "unit": "ns",
            "range": "± 6323435.865085131"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "gabrielandremarquez.matte@gmail.com",
            "name": "Gabriel Matte",
            "username": "GabrielMarquezMatte"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "cf80f86b57ea36b0d93d316f8bf35ac7449d8cf4",
          "message": "Merge pull request #8 from GabrielMarquezMatte/develop\n\nEnhance XLS file handling with tests, refactoring, and new writers",
          "timestamp": "2026-06-25T08:20:53-03:00",
          "tree_id": "6542cf68fe17d847ed68a22aacaeb586de58a7f7",
          "url": "https://github.com/GabrielMarquezMatte/ExcelReader/commit/cf80f86b57ea36b0d93d316f8bf35ac7449d8cf4"
        },
        "date": 1782386570660,
        "tool": "benchmarkdotnet",
        "benches": [
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserSync(Rows: 50000)",
            "value": 40767698.24786324,
            "unit": "ns",
            "range": "± 114790.2053892417"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserAsync(Rows: 50000)",
            "value": 46804384.03409091,
            "unit": "ns",
            "range": "± 28619.816596977806"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.MiniExcel(Rows: 50000)",
            "value": 354616932.2222222,
            "unit": "ns",
            "range": "± 848225.5343285442"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.Sylvan(Rows: 50000)",
            "value": 108525734.66666667,
            "unit": "ns",
            "range": "± 404974.28624018835"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.SylvanAsync(Rows: 50000)",
            "value": 114345110.8888889,
            "unit": "ns",
            "range": "± 1419967.2197747598"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReader(Rows: 50000)",
            "value": 37675734.690476194,
            "unit": "ns",
            "range": "± 128272.68469545941"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReaderAsync(Rows: 50000)",
            "value": 42157753.78703704,
            "unit": "ns",
            "range": "± 109344.98039361446"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.MiniExcel(Rows: 50000)",
            "value": 342998570.5,
            "unit": "ns",
            "range": "± 1972743.0795337998"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.Sylvan(Rows: 50000)",
            "value": 74454848.37142858,
            "unit": "ns",
            "range": "± 371294.71978835383"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.ExcelReaderWriter(Rows: 50000)",
            "value": 47414269.8,
            "unit": "ns",
            "range": "± 100259.43932735958"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.MiniExcel(Rows: 50000)",
            "value": 108680891.2,
            "unit": "ns",
            "range": "± 5825394.85883117"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsReadBenchmark.ExcelReader(Rows: 50000)",
            "value": 6909349.284722222,
            "unit": "ns",
            "range": "± 30994.08633649651"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsReadBenchmark.ExcelReaderAsync(Rows: 50000)",
            "value": 6827966.0046875,
            "unit": "ns",
            "range": "± 13553.050716993659"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsReadBenchmark.Sylvan(Rows: 50000)",
            "value": 8579712.6171875,
            "unit": "ns",
            "range": "± 45900.34374005789"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsWriteBenchmark.XlsWriter(Rows: 50000)",
            "value": 12389376.927083334,
            "unit": "ns",
            "range": "± 161244.7126831152"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsWriteBenchmark.XlsxWriter(Rows: 50000)",
            "value": 49744328.236363634,
            "unit": "ns",
            "range": "± 205788.08606495068"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "gabrielandremarquez.matte@gmail.com",
            "name": "Gabriel Matte",
            "username": "GabrielMarquezMatte"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "1a96c507d6afc00ebe99c84ce465f628af0265b9",
          "message": "Merge pull request #9 from GabrielMarquezMatte/develop\n\nEnhance XLS writer with sheet splitting and multi-framework support",
          "timestamp": "2026-06-25T18:02:26-03:00",
          "tree_id": "6a3ebc1049c4ef3e046420f3e5103ef7d3b35ef6",
          "url": "https://github.com/GabrielMarquezMatte/ExcelReader/commit/1a96c507d6afc00ebe99c84ce465f628af0265b9"
        },
        "date": 1782421457329,
        "tool": "benchmarkdotnet",
        "benches": [
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserSync(Rows: 50000)",
            "value": 40376161.39230769,
            "unit": "ns",
            "range": "± 212890.3639249534"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserAsync(Rows: 50000)",
            "value": 46538311.38383838,
            "unit": "ns",
            "range": "± 374406.32315018494"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.MiniExcel(Rows: 50000)",
            "value": 351386749.2,
            "unit": "ns",
            "range": "± 2347018.1276029944"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.Sylvan(Rows: 50000)",
            "value": 112220804.55555555,
            "unit": "ns",
            "range": "± 752284.8929150963"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.SylvanAsync(Rows: 50000)",
            "value": 114316159.44444445,
            "unit": "ns",
            "range": "± 1142364.6547697796"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReader(Rows: 50000)",
            "value": 37923694.12857143,
            "unit": "ns",
            "range": "± 443501.2287526622"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReaderAsync(Rows: 50000)",
            "value": 41728375.0462963,
            "unit": "ns",
            "range": "± 113307.92505535137"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.MiniExcel(Rows: 50000)",
            "value": 321409144.2,
            "unit": "ns",
            "range": "± 1049006.310097619"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.Sylvan(Rows: 50000)",
            "value": 78009967.34285714,
            "unit": "ns",
            "range": "± 473241.19578184315"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.ExcelReaderWriter(Rows: 50000)",
            "value": 49813401.233333334,
            "unit": "ns",
            "range": "± 102772.94330075393"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.MiniExcel(Rows: 50000)",
            "value": 104414334.8,
            "unit": "ns",
            "range": "± 1594985.5429498577"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsReadBenchmark.ExcelReader(Rows: 50000)",
            "value": 6759899.7484375,
            "unit": "ns",
            "range": "± 42915.49580766053"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsReadBenchmark.ExcelReaderAsync(Rows: 50000)",
            "value": 6672800.974826389,
            "unit": "ns",
            "range": "± 20136.587571038828"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsReadBenchmark.Sylvan(Rows: 50000)",
            "value": 8570584.571875,
            "unit": "ns",
            "range": "± 33248.922692395194"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsWriteBenchmark.XlsWriter(Rows: 50000)",
            "value": 12121261.065625,
            "unit": "ns",
            "range": "± 137897.82122605765"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsWriteBenchmark.XlsxWriter(Rows: 50000)",
            "value": 48210594.464646466,
            "unit": "ns",
            "range": "± 340682.40217490937"
          }
        ]
      },
      {
        "commit": {
          "author": {
            "email": "gabrielandremarquez.matte@gmail.com",
            "name": "Gabriel Matte",
            "username": "GabrielMarquezMatte"
          },
          "committer": {
            "email": "noreply@github.com",
            "name": "GitHub",
            "username": "web-flow"
          },
          "distinct": true,
          "id": "2ce805547ba13619ed1edad4ccbdceba91001e46",
          "message": "Merge pull request #10 from GabrielMarquezMatte/develop",
          "timestamp": "2026-06-26T02:44:13-03:00",
          "tree_id": "237abce2899264defb35080aef1029290353ed42",
          "url": "https://github.com/GabrielMarquezMatte/ExcelReader/commit/2ce805547ba13619ed1edad4ccbdceba91001e46"
        },
        "date": 1782452801736,
        "tool": "benchmarkdotnet",
        "benches": [
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserSync(Rows: 50000)",
            "value": 42420277.395833336,
            "unit": "ns",
            "range": "± 113012.47510231589"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserAsync(Rows: 50000)",
            "value": 46968938.31818181,
            "unit": "ns",
            "range": "± 20609.614404867876"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserXlsbSync(Rows: 50000)",
            "value": 13219861.115625,
            "unit": "ns",
            "range": "± 20186.086339509195"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.ExcelParserXlsbAsync(Rows: 50000)",
            "value": 15531083.354166666,
            "unit": "ns",
            "range": "± 55432.02522091011"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.MiniExcel(Rows: 50000)",
            "value": 338280932.4,
            "unit": "ns",
            "range": "± 3532326.9419557853"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.Sylvan(Rows: 50000)",
            "value": 107255144.875,
            "unit": "ns",
            "range": "± 154848.58442679464"
          },
          {
            "name": "ExcelReader.Benchmarks.ParseBenchmark.SylvanAsync(Rows: 50000)",
            "value": 115594881.8,
            "unit": "ns",
            "range": "± 5634747.457615722"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReader(Rows: 50000)",
            "value": 37566646.60714286,
            "unit": "ns",
            "range": "± 33769.21970173176"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReaderAsync(Rows: 50000)",
            "value": 42220749.10185186,
            "unit": "ns",
            "range": "± 41845.72013812361"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReaderXlsb(Rows: 50000)",
            "value": 5443956.578125,
            "unit": "ns",
            "range": "± 12454.802364655448"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.ExcelReaderXlsbAsync(Rows: 50000)",
            "value": 6499375.3974609375,
            "unit": "ns",
            "range": "± 6455.81879377307"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.MiniExcel(Rows: 50000)",
            "value": 321595597.3333333,
            "unit": "ns",
            "range": "± 2102157.4772704947"
          },
          {
            "name": "ExcelReader.Benchmarks.ReadBenchmark.Sylvan(Rows: 50000)",
            "value": 71727205.47142856,
            "unit": "ns",
            "range": "± 694222.871675797"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.ExcelReaderWriter(Rows: 50000)",
            "value": 47631295.7070707,
            "unit": "ns",
            "range": "± 408158.121076346"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.ExcelReaderXlsbWriter(Rows: 50000)",
            "value": 16643721.578125,
            "unit": "ns",
            "range": "± 63921.80354151172"
          },
          {
            "name": "ExcelReader.Benchmarks.WriteBenchmark.MiniExcel(Rows: 50000)",
            "value": 102784908.13333336,
            "unit": "ns",
            "range": "± 1835243.7386273781"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsReadBenchmark.ExcelReader(Rows: 50000)",
            "value": 6750072.6015625,
            "unit": "ns",
            "range": "± 21973.088626242396"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsReadBenchmark.ExcelReaderAsync(Rows: 50000)",
            "value": 6678877.9109375,
            "unit": "ns",
            "range": "± 12630.185548789694"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsReadBenchmark.Sylvan(Rows: 50000)",
            "value": 8553921.8859375,
            "unit": "ns",
            "range": "± 27066.230052941562"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsWriteBenchmark.XlsWriter(Rows: 50000)",
            "value": 12337732.8578125,
            "unit": "ns",
            "range": "± 139908.36350960605"
          },
          {
            "name": "ExcelReader.Benchmarks.XlsWriteBenchmark.XlsxWriter(Rows: 50000)",
            "value": 48679975.36363636,
            "unit": "ns",
            "range": "± 161431.13633437257"
          }
        ]
      }
    ]
  }
}