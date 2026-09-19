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
    } 

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
        }

    private:
        int64_t start_allocs_ = 0;
        int64_t start_bytes_ = 0;
    };

} 

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
