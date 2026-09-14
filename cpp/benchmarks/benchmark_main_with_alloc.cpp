// Replaces linking benchmark::benchmark_main: RegisterMemoryManager must run before
// RunSpecifiedBenchmarks(), which benchmark_main's own canned main() gives no hook for.
//
// Allocation metrics only reach the JSON reporter (see Google Benchmark's docs/user_guide.md,
// "Memory Usage" — "This data will then be reported alongside other performance data, currently
// only when using JSON output."), so console-only runs never show them. Run with, e.g.:
//   excelreader_cpp_benchmarks.exe --benchmark_out=results.json --benchmark_out_format=json
// `allocs_per_iter`/`total_allocated_bytes` land in results.json per benchmark, on top of whatever
// --benchmark_format=console still prints to the terminal.
#include "alloc_memory_manager.hpp"

#include <benchmark/benchmark.h>

int main(int argc, char **argv)
{
    xl_bench::AllocMemoryManager memory_manager;
    benchmark::RegisterMemoryManager(&memory_manager);

    benchmark::Initialize(&argc, argv);
    if (benchmark::ReportUnrecognizedArguments(argc, argv))
    {
        return 1;
    }
    benchmark::RunSpecifiedBenchmarks();
    benchmark::RegisterMemoryManager(nullptr);
    benchmark::Shutdown();
    return 0;
}
