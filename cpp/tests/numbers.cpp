// Built twice: once as is, and once with EXCELREADER_TEST_NO_TO_CHARS, which hides the standard
// library's floating-point <charconv> to exercise the fallback a libc++ without it compiles.
#include <charconv>
#ifdef EXCELREADER_TEST_NO_TO_CHARS
#undef __cpp_lib_to_chars
#endif
#include <xl/excelreader.hpp>

#include <cstdio>
#include <cstdlib>
#include <limits>
#include <string>

namespace
{
    int failures = 0;

    void check(bool condition, const char *what)
    {
        if (!condition)
        {
            std::fprintf(stderr, "FAILED: %s\n", what);
            ++failures;
        }
    }

    xl::CellView text_cell(std::string_view text)
    {
        return xl::CellView{0, xl::CellType::Number, text, std::nullopt};
    }

    void round_trips(double value)
    {
        const std::string text = xl::CellView{0, xl::CellType::Number, {}, value}.text();
        check(text_cell(text).as_number() == value, "text() of a number parses back to the same double");
    }
}

int main()
{
    check(text_cell("12.5").as_number() == 12.5, "a decimal parses");
    check(text_cell("1e3").as_number() == 1000.0, "an exponent parses");
    check(text_cell("-7").as_number() == -7.0, "a negative parses");
    check(!text_cell("abc").as_number().has_value(), "text is not a number");
    check(!text_cell("").as_number().has_value(), "empty is not a number");
    check(!text_cell(" 5").as_number().has_value(), "leading space is not a number");
    check(!text_cell("5x").as_number().has_value(), "trailing garbage is not a number");
    check(!text_cell("1,5").as_number().has_value(), "a comma decimal is not a number");

    check(xl::CellView{0, xl::CellType::Number, {}, 42.5}.text() == "42.5", "a short number formats short");
    check(xl::CellView{0, xl::CellType::Number, {}, 3.0}.text() == "3", "an integral number has no fraction");
    check(xl::CellView{0, xl::CellType::Number, "kept", 1.0}.text() == "kept", "stored text wins over the number");
    check(xl::CellView{0, xl::CellType::Number, {}, 1000.0}.text() == "1000", "fixed wins when it is shorter");
    check(xl::CellView{0, xl::CellType::Number, {}, 0.001}.text() == "0.001", "fixed wins a tie");
    check(xl::CellView{0, xl::CellType::Number, {}, 1e20}.text() == "1e+20", "scientific wins when it is shorter");
    check(xl::CellView{0, xl::CellType::Number, {}, 1e-7}.text() == "1e-07", "a small number goes scientific");
    check(xl::CellView{0, xl::CellType::Number, {}, 123456789012345680000.0}.text() == "123456789012345683968",
          "a large integer prints its exact value");
    check(xl::CellView{0, xl::CellType::Number, {}, 0.0001}.text() == "1e-04", "scientific wins at 1e-04");
    check(xl::CellView{0, xl::CellType::Number, {}, 123456.0}.text() == "123456", "an integer stays fixed");
    check(xl::CellView{0, xl::CellType::Number, {}, 1e15}.text() == "1e+15", "a round power of ten goes scientific");
    check(xl::CellView{0, xl::CellType::Number, {}, 0.0125}.text() == "0.0125", "a fraction places its point");
    check(xl::CellView{0, xl::CellType::Number, {}, -2.5}.text() == "-2.5", "a negative keeps its sign");
    check(xl::CellView{0, xl::CellType::Number, {}, 0.0}.text() == "0", "zero is 0");
    check(xl::CellView{0, xl::CellType::Number, {}, std::numeric_limits<double>::infinity()}.text() == "inf",
          "infinity has no exponent to place");
    round_trips(0.1);
    round_trips(1.0 / 3.0);
    round_trips(1e21);
    round_trips(-2.5e-300);
    round_trips(45000.25);

    if (failures != 0)
    {
        std::fprintf(stderr, "%d check(s) failed\n", failures);
        return EXIT_FAILURE;
    }
    std::puts("numbers: OK");
    return EXIT_SUCCESS;
}
