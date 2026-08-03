namespace ExcelReader.CodSpeed
{
    // One CodSpeed benchmark: a named workload plus the number of times it is repeated inside a
    // single process.
    //
    // CodSpeed's CLI harness times the whole process, so every scenario pays a constant cost for
    // runtime startup, JIT and fixture loading. `Iterations` is tuned so the measured workload
    // dominates that offset while each process still finishes in a few hundred milliseconds.
    internal sealed class Scenario
    {
        private readonly Func<int, Task<long>> _run;

        public Scenario(string name, int iterations, Func<int, Task<long>> run)
        {
            Name = name;
            Iterations = iterations;
            _run = run;
        }

        public string Name { get; }

        public int Iterations { get; }

        public Task<long> RunAsync(int iterations) => _run(iterations);
    }
}
