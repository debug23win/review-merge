using System;
using System.Linq;
using System.Threading;
using ReviewMerge;

static class BoxAssignmentsTest {
    static int Main(string[] args) {
        int prompts = 0;
        string mode = args[1];
        var options = new MergeOptions { RecalculateHistory = true };
        options.ChooseBoxReviewer = conflict => {
            prompts++;
            if (conflict.Box != "ПД №11" || !conflict.Reviewers.Select(r => r.Name + ":" + r.Files).SequenceEqual(new[] { "Борис:2", "Анна:1" }))
                throw new Exception("Incorrect participants or file counts for shared box");
            if (conflict.Recommended == null || conflict.Recommended.Name != "Борис")
                throw new Exception("Reviewer with the most files should be recommended");
            if (mode == "cancel") return null;
            return mode == "override" ? "Анна" : conflict.Recommended.Name;
        };
        try {
            var result = MergeEngine.Run(new string[0], args[0], "Все загруженные файлы", new DateTime(2026, 10, 4), (p, line) => {}, CancellationToken.None, options);
            if (mode == "cancel" || result.BoxAssignments.Count != 1)
                throw new Exception("Box assignment was not recorded or cancellation was ignored");
        } catch (OperationCanceledException) {
            if (mode != "cancel") throw;
        }
        if (prompts != 1) throw new Exception("Every build must ask once for the shared box");
        return 0;
    }
}
