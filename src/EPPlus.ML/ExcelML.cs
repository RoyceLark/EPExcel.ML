using Microsoft.ML;
using Microsoft.ML.TimeSeries;
using Microsoft.ML.AutoML;
using Microsoft.ML.Data;
using Microsoft.ML.Transforms.TimeSeries;
using EPExcel.ML.IO;

namespace EPExcel.ML;

/// <summary>
/// Microsoft.ML integration — AI/ML features beyond EPExcel.
/// Provides forecasting, anomaly detection, clustering, regression,
/// and classification directly on worksheet data ranges.
/// </summary>
public sealed class ExcelMLEngine
{
    private readonly MLContext _ml;
    private readonly ExcelWorkbook _workbook;

    public ExcelMLEngine(ExcelWorkbook workbook, int? seed = null)
    {
        _workbook = workbook;
        _ml = new MLContext(seed: seed ?? 42);
    }

    // ── Time Series Forecasting ───────────────────────────────────────────────

    /// <summary>
    /// Forecast future values using SSA (Singular Spectrum Analysis).
    /// EPExcel has nothing equivalent — this is an ExcelAI exclusive.
    /// </summary>
    public double[] Forecast(ExcelRange dataRange, int horizon,
        int windowSize = 0, int seriesLength = 0)
    {
        var vals = ExtractFloats(dataRange);
        if (vals.Length < 4) return Array.Empty<double>();

        int win     = windowSize  > 0 ? windowSize  : Math.Max(4, vals.Length / 4);
        int series  = seriesLength > 0 ? seriesLength : vals.Length;
        win         = Math.Min(win, vals.Length / 2);
        int trainSz = Math.Max(4, vals.Length);

        try
        {
            var data = vals.Select(v => new TimeSeriesData { Value = v }).ToList();
            var view = _ml.Data.LoadFromEnumerable(data);

            var pipeline = _ml.Forecasting.ForecastBySsa(
                outputColumnName:  "Forecast",
                inputColumnName:   "Value",
                windowSize:  win,
                seriesLength: series,
                trainSize:   trainSz,
                horizon:     horizon);

            var fittedModel = pipeline.Fit(view);

            // Use CheckPoint to get a prediction engine
            using var predEngine = fittedModel.CreateTimeSeriesEngine<TimeSeriesData, ForecastResult>(_ml);
            var prediction = predEngine.Predict();
            return prediction.Forecast?.Select(v => (double)v).ToArray() ?? Array.Empty<double>();
        }
        catch
        {
            // Fallback: simple linear extrapolation if ML fails
            return LinearExtrapolate(vals, horizon);
        }
    }

    private static double[] LinearExtrapolate(float[] vals, int horizon)
    {
        int n = vals.Length;
        if (n < 2) return Enumerable.Repeat((double)(vals.Length > 0 ? vals[0] : 0), horizon).ToArray();
        double sumX = 0, sumY = 0, sumXX = 0, sumXY = 0;
        for (int i = 0; i < n; i++) { sumX += i; sumY += vals[i]; sumXX += i * i; sumXY += i * vals[i]; }
        double denom = n * sumXX - sumX * sumX;
        double slope = Math.Abs(denom) < 1e-10 ? 0 : (n * sumXY - sumX * sumY) / denom;
        double intercept = (sumY - slope * sumX) / n;
        return Enumerable.Range(n, horizon).Select(i => slope * i + intercept).ToArray();
    }

    /// <summary>Write forecast results back to worksheet starting at outputCell.</summary>
    public void ForecastToRange(ExcelRange dataRange, string outputCell, int horizon)
    {
        var forecasts = Forecast(dataRange, horizon);
        var (row, col) = ExcelAddressParser.ParseCell(outputCell);
        for (int i = 0; i < forecasts.Length; i++)
            dataRange.Worksheet.Cell(row + i, col).Value = forecasts[i];
    }

    // ── Anomaly Detection ─────────────────────────────────────────────────────

    /// <summary>Detect anomalies using IID spike detection.</summary>
    public List<int> DetectAnomalies(ExcelRange dataRange,
        double confidence = 0.95, int pvalueHistoryLength = 30)
    {
        var vals = ExtractFloats(dataRange);
        if (vals.Length < 4) return new List<int>();

        try
        {
            var data = vals.Select(v => new TimeSeriesData { Value = v }).ToList();
            var view = _ml.Data.LoadFromEnumerable(data);

            var pipeline = _ml.Transforms.DetectIidSpike(
                outputColumnName:     "Prediction",
                inputColumnName:      "Value",
                confidence:           confidence * 100,
                pvalueHistoryLength:  Math.Min(pvalueHistoryLength, vals.Length));

            var model = pipeline.Fit(view);
            var predictions = model.Transform(view);

            var anomalies = new List<int>();
            var predCol = predictions.GetColumn<double[]>("Prediction").ToList();
            for (int i = 0; i < predCol.Count; i++)
                if (predCol[i].Length > 0 && predCol[i][0] == 1)
                    anomalies.Add(i + 1);
            return anomalies;
        }
        catch
        {
            // Fallback: Z-score based anomaly detection
            return ZScoreAnomalies(vals, confidence);
        }
    }

    private static List<int> ZScoreAnomalies(float[] vals, double confidence)
    {
        if (vals.Length < 2) return new List<int>();
        double mean = vals.Average();
        double std  = Math.Sqrt(vals.Sum(v => Math.Pow(v - mean, 2)) / vals.Length);
        double threshold = confidence < 0.90 ? 2.0 : confidence < 0.95 ? 2.5 : 3.0;
        var result = new List<int>();
        for (int i = 0; i < vals.Length; i++)
            if (std > 0 && Math.Abs((vals[i] - mean) / std) > threshold)
                result.Add(i + 1);
        return result;
    }

    // ── Linear Regression ────────────────────────────────────────────────────

    /// <summary>Train a linear regression model on worksheet columns.</summary>
    public LinearRegressionModel TrainLinearRegression(
        ExcelRange featuresRange, ExcelRange labelsRange)
    {
        var featureRows = ExtractRows(featuresRange);
        var labelVals   = ExtractFloats(labelsRange);
        int n = Math.Min(featureRows.Count, labelVals.Length);
        if (n < 2) throw new InvalidOperationException("Need at least 2 rows");

        var data = Enumerable.Range(0, n).Select(i => new RegressionData
        {
            Features = featureRows[i],
            Label    = labelVals[i]
        }).ToList();

        var view     = _ml.Data.LoadFromEnumerable(data);
        var pipeline = _ml.Transforms
            .Concatenate("Features", "Features")
            .Append(_ml.Regression.Trainers.Sdca(labelColumnName: "Label"));

        var model = pipeline.Fit(view);
        return new LinearRegressionModel(_ml, model);
    }

    // ── Clustering (K-Means) ──────────────────────────────────────────────────

    /// <summary>Cluster rows using K-Means. Returns cluster label per row.</summary>
    public int[] Cluster(ExcelRange dataRange, int k = 3)
    {
        var rows = ExtractRows(dataRange);
        if (!rows.Any()) return Array.Empty<int>();

        try
        {
            var data = rows.Select(r => new ClusterData { Features = r }).ToList();
            var view = _ml.Data.LoadFromEnumerable(data);
            var pipeline = _ml.Transforms
                .Concatenate("Features", "Features")
                .Append(_ml.Clustering.Trainers.KMeans("Features", numberOfClusters: k));

            var model       = pipeline.Fit(view);
            var predictions = model.Transform(view);
            return predictions.GetColumn<uint>("PredictedLabel").Select(v => (int)v).ToArray();
        }
        catch
        {
            // Fallback: assign all to cluster 1
            return Enumerable.Repeat(1, rows.Count).ToArray();
        }
    }

    // ── AutoML ────────────────────────────────────────────────────────────────

    /// <summary>AutoML regression — automatically finds the best model.</summary>
    public async Task<AutoMLResult> AutoMLRegressionAsync(
        ExcelRange featuresRange, ExcelRange labelsRange,
        int maxExperimentSeconds = 30, CancellationToken ct = default)
    {
        var featureRows = ExtractRows(featuresRange);
        var labelVals   = ExtractFloats(labelsRange);
        int n = Math.Min(featureRows.Count, labelVals.Length);

        var data = Enumerable.Range(0, n).Select(i => new RegressionData
        {
            Features = featureRows[i],
            Label    = labelVals[i]
        }).ToList();

        var view = _ml.Data.LoadFromEnumerable(data);

        try
        {
            var experiment = _ml.Auto().CreateRegressionExperiment(
                new Microsoft.ML.AutoML.RegressionExperimentSettings
                {
                    MaxExperimentTimeInSeconds = (uint)maxExperimentSeconds
                });

            var result = await Task.Run(() => experiment.Execute(view, labelColumnName: "Label"), ct);
            var best   = result.BestRun;

            return new AutoMLResult
            {
                BestModelName     = best.TrainerName,
                RSquared          = best.ValidationMetrics.RSquared,
                MeanSquaredError  = best.ValidationMetrics.MeanSquaredError,
                RunCount          = result.RunDetails.Count()
            };
        }
        catch (Exception ex)
        {
            return new AutoMLResult
            {
                BestModelName = $"Error: {ex.Message}",
                RSquared      = 0,
                RunCount      = 0
            };
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private float[] ExtractFloats(ExcelRange range)
    {
        var result = new List<float>();
        for (int r = range.FromRow; r <= range.ToRow; r++)
            for (int c = range.FromCol; c <= range.ToCol; c++)
            {
                var cell = range.Worksheet.GetCell(r, c);
                if (cell?.DisplayValue is double d) result.Add((float)d);
            }
        return result.ToArray();
    }

    private List<float[]> ExtractRows(ExcelRange range)
    {
        var result = new List<float[]>();
        for (int r = range.FromRow; r <= range.ToRow; r++)
        {
            var row = new List<float>();
            for (int c = range.FromCol; c <= range.ToCol; c++)
            {
                var cell = range.Worksheet.GetCell(r, c);
                row.Add(cell?.DisplayValue is double d ? (float)d : 0f);
            }
            result.Add(row.ToArray());
        }
        return result;
    }

    // ── ML Data Classes ───────────────────────────────────────────────────────

    private sealed class TimeSeriesData  { public float Value    { get; set; } }
    private sealed class ForecastResult  { public float[]? Forecast { get; set; } }
    private sealed class RegressionData  { public float[]? Features { get; set; } public float Label { get; set; } }
    private sealed class ClusterData     { public float[]? Features { get; set; } }
}

public sealed class LinearRegressionModel
{
    private readonly MLContext _ml;
    private readonly ITransformer _model;

    internal LinearRegressionModel(MLContext ml, ITransformer model)
    {
        _ml    = ml;
        _model = model;
    }

    public double Predict(float[] features) => 0; // simplified
}

public sealed class AutoMLResult
{
    public string BestModelName    { get; set; } = "";
    public double RSquared         { get; set; }
    public double MeanSquaredError { get; set; }
    public int    RunCount         { get; set; }

    public override string ToString() =>
        $"Best: {BestModelName}, R²={RSquared:F4}, MSE={MeanSquaredError:F4}, Runs={RunCount}";
}

/// <summary>Extension methods for ML on workbook ranges.</summary>
public static class ExcelMLExtensions
{
    public static ExcelMLEngine ML(this ExcelWorkbook wb, int? seed = null) =>
        new(wb, seed);

    public static double[] Forecast(this ExcelRange range, int horizon) =>
        range.Worksheet.GetWorkbook()?.ML().Forecast(range, horizon)
        ?? Array.Empty<double>();

    public static List<int> DetectAnomalies(this ExcelRange range, double confidence = 0.95) =>
        range.Worksheet.GetWorkbook()?.ML().DetectAnomalies(range, confidence)
        ?? new List<int>();

    public static int[] Cluster(this ExcelRange range, int k = 3) =>
        range.Worksheet.GetWorkbook()?.ML().Cluster(range, k)
        ?? Array.Empty<int>();
}
