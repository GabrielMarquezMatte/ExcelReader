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
