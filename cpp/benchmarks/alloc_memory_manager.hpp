// Tracks C++-heap allocation activity between benchmark::MemoryManager::Start()/Stop(), via global
// operator new/delete overrides. This only sees allocations made through the C++ runtime's own
// `new` — the columnar buffers xl_parse_typed/xl_free_table hand across the FFI boundary are
// allocated (and freed) entirely inside the NativeAOT library, on its own GC heap, and never touch
// this process's operator new. What DOES show up here: every std::string/std::vector construction
// xl::TableView<T>'s row materialization and xl::Workbook::open's own bookkeeping perform — the same
// scope BenchmarkDotNet's MemoryDiagnoser reports for the C# benchmarks (managed-heap allocations,
// not the unmanaged buffers underneath them). The two numbers are comparable on that basis, not on
// an "everything the process touched" basis.
#pragma once

#include <atomic>
#include <benchmark/benchmark.h>
#include <cstdlib>
#include <new>

namespace xl_bench
{
    namespace detail
    {
        inline std::atomic<int64_t> g_num_allocs{0};
        inline std::atomic<int64_t> g_bytes_allocated{0};
    } // namespace detail

    // Registered once, in main() — see benchmark_main.cpp. Start()/Stop() snapshot the global
    // counters rather than reset them, so overlapping/nested benchmark iterations (there are none
    // here, but the counters are process-wide) never lose a count to a reset race.
    class AllocMemoryManager : public benchmark::MemoryManager
    {
    public:
        void Start() override
        {
            start_allocs_ = detail::g_num_allocs.load(std::memory_order_relaxed);
            start_bytes_ = detail::g_bytes_allocated.load(std::memory_order_relaxed);
        }

        void Stop(Result &result) override
        {
            result.num_allocs = detail::g_num_allocs.load(std::memory_order_relaxed) - start_allocs_;
            result.total_allocated_bytes =
                detail::g_bytes_allocated.load(std::memory_order_relaxed) - start_bytes_;
            // max_bytes_used/net_heap_growth would need matching operator delete instrumentation
            // (sized delete, to know how much was freed) — not tracked here. total_allocated_bytes
            // alone already answers "how much did this allocate", the same question
            // BenchmarkDotNet's Allocated column answers; peak/net usage is a different question
            // this repo has not needed answered yet.
        }

    private:
        int64_t start_allocs_ = 0;
        int64_t start_bytes_ = 0;
    };

} // namespace xl_bench

// Every standard scalar/array, throwing/nothrow new overload is counted the same way; delete is left
// at its default (malloc/free) — freeing doesn't need instrumentation to answer "bytes allocated".
void *operator new(std::size_t size)
{
    void *ptr = std::malloc(size != 0 ? size : 1);
    if (ptr == nullptr)
    {
        throw std::bad_alloc();
    }
    xl_bench::detail::g_num_allocs.fetch_add(1, std::memory_order_relaxed);
    xl_bench::detail::g_bytes_allocated.fetch_add(static_cast<int64_t>(size), std::memory_order_relaxed);
    return ptr;
}

void *operator new(std::size_t size, const std::nothrow_t &) noexcept
{
    void *ptr = std::malloc(size != 0 ? size : 1);
    if (ptr != nullptr)
    {
        xl_bench::detail::g_num_allocs.fetch_add(1, std::memory_order_relaxed);
        xl_bench::detail::g_bytes_allocated.fetch_add(static_cast<int64_t>(size), std::memory_order_relaxed);
    }
    return ptr;
}

void *operator new[](std::size_t size)
{
    return ::operator new(size);
}

void *operator new[](std::size_t size, const std::nothrow_t &tag) noexcept
{
    return ::operator new(size, tag);
}

void operator delete(void *ptr) noexcept
{
    std::free(ptr);
}

void operator delete(void *ptr, std::size_t) noexcept
{
    std::free(ptr);
}

void operator delete[](void *ptr) noexcept
{
    std::free(ptr);
}

void operator delete[](void *ptr, std::size_t) noexcept
{
    std::free(ptr);
}
