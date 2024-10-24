using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.ML;
using Microsoft.ML.Data;
using Newtonsoft.Json;
using System.Text;
using System.Net.Http;
using System.Threading.Tasks;

namespace SentimentAnalysis
{
    // Input data class (text for sentiment analysis)
    public class SentimentData
    {
        [LoadColumn(0)]
        public string SentimentText;

        [LoadColumn(1)]
        public int Sentiment;  // 0 for Negative, 1 for Neutral, 2 for Positive
    }

    // Prediction output class
    public class SentimentPrediction : SentimentData
    {
        [ColumnName("PredictedLabel")]
        public int PredictedSentiment;
        public float[] Score;
    }

    internal class Program
    {        
        static async Task Main(string[] args)
        {
            SentimentAnalysisTByTrainingData();            
        }

        private static void SentimentAnalysisTByTrainingData()
        {
            // Create ML context
            MLContext mlContext = new MLContext();
            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var trainingDatafilePath = Path.Combine(baseDirectory, "training_data.xlsx");// @"C:\Users\dushyantchouhan\source\repos\SentimentAnalysis\training_data.xlsx"; 
            var inputFilePath = Path.Combine(baseDirectory, "training_data.xlsx");// @"C:\Users\dushyantchouhan\source\repos\SentimentAnalysis\employee_reviews1.xlsx";
            var outputFilePath = Path.Combine(baseDirectory, "training_data.xlsx");// @"C:\Users\dushyantchouhan\source\repos\SentimentAnalysis\prediction_results.xlsx";

            if (!File.Exists(trainingDatafilePath))
            {
                Console.WriteLine($"File not found: {trainingDatafilePath}");
                return;
            }

            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"Input file not found: {inputFilePath}");
                return;
            }

            var data = LoadDataFromExcel(trainingDatafilePath);
            // Convert list to IDataView
            IDataView trainingData = mlContext.Data.LoadFromEnumerable(data);

            // Create a pipeline for training the model
            var pipeline = mlContext.Transforms.Text.FeaturizeText("Features", nameof(SentimentData.SentimentText))
                            .Append(mlContext.Transforms.Conversion.MapValueToKey("Label", nameof(SentimentData.Sentiment)))
                            .Append(mlContext.Transforms.Concatenate("Features", "Features"))
                            .Append(mlContext.MulticlassClassification.Trainers.SdcaMaximumEntropy("Label", "Features"))
                            .Append(mlContext.Transforms.Conversion.MapKeyToValue("PredictedLabel"));

            // Train the model
            var model = pipeline.Fit(trainingData);

            // Load prediction data from Excel
            var predictionData = LoadMultipleDataFromExcel(inputFilePath);

            // Convert input data to IDataView for prediction
            var predictionEngine = mlContext.Model.CreatePredictionEngine<SentimentData, SentimentPrediction>(model);

            var results = new List<SentimentPrediction>();

            for (int i = 0; i < predictionData.Count; i++)
            {
                var inputData = predictionData[i];
                var prediction = predictionEngine.Predict(inputData);
                results.Add(prediction);
            }

            // Save results to Excel
            SaveResultsToExcel(results, outputFilePath);

        }

        private static List<SentimentData> LoadDataFromExcel(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var data = new List<SentimentData>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var table = result.Tables[0];

                    for (int i = 1; i < table.Rows.Count; i++) // Start from 1 to skip header row
                    {
                        var row = table.Rows[i];
                        data.Add(new SentimentData
                        {
                            SentimentText = row[0].ToString(),
                            Sentiment = MapSentiment(row[1].ToString().ToLower()) // Assuming the sentiment is stored as 0, 1, or 2
                        });
                    }
                }
            }

            return data;
        }

        private static int MapSentiment(string sentiment)
        {
            return sentiment switch
            {
                "negative" => 0,
                "neutral" => 1,
                "positive" => 2,
                _ => throw new ArgumentException("Invalid sentiment value")
            };
        }

        private static string GetSentimentLabel(int sentiment)
        {
            return sentiment switch
            {
                0 => "Negative",
                1 => "Neutral",
                2 => "Positive",
                _ => "Unknown"
            };
        }

        private static List<SentimentData> LoadMultipleDataFromExcel(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var data = new List<SentimentData>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    var table = result.Tables[0];

                    for (int i = 1; i < table.Rows.Count; i++) // Start from 1 to skip header row
                    {
                        var row = table.Rows[i];
                        data.Add(new SentimentData
                        {
                            SentimentText = row[0].ToString(),
                            Sentiment = -1 // Default value since sentiment is not provided
                        });
                    }
                }
            }

            return data;
        }

        private static void SaveResultsToExcel(List<SentimentPrediction> results, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Predictions");
                worksheet.Cell(1, 1).Value = "Text";
                worksheet.Cell(1, 2).Value = "Predicted Sentiment";
                worksheet.Cell(1, 3).Value = "Scores";

                for (int i = 0; i < results.Count; i++)
                {
                    var result = results[i];
                    worksheet.Cell(i + 2, 1).Value = result.SentimentText;
                    worksheet.Cell(i + 2, 2).Value = GetSentimentLabel(result.PredictedSentiment);
                    worksheet.Cell(i + 2, 3).Value = string.Join(", ", result.Score);
                }

                workbook.SaveAs(filePath);
            }
        }

        //private async static void SentimentAnalysisByChatGpt()
        //{
        //    var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        //    var trainingDatafilePath = Path.Combine(baseDirectory, "training_data.xlsx");// @"C:\Users\dushyantchouhan\source\repos\SentimentAnalysis\training_data.xlsx"; 
        //    var inputFilePath = Path.Combine(baseDirectory, "training_data.xlsx");// @"C:\Users\dushyantchouhan\source\repos\SentimentAnalysis\employee_reviews1.xlsx";
        //    var outputFilePath = Path.Combine(baseDirectory, "training_data.xlsx");// @"C:\Users\dushyantchouhan\source\repos\SentimentAnalysis\prediction_results.xlsx";
        //    var openAiApiKey = ""; // enter chat gpt API
        //    // Load prediction data from Excel
        //    var predictionData = LoadMultipleDataFromExcel(inputFilePath);

        //    // Initialize OpenAI client
        //    var openAiClient = new OpenAIClient(openAiApiKey);

        //    var results = new List<SentimentPrediction>();

        //    for (int i = 0; i < 10; i++)
        //    {
        //        var inputData = predictionData[i];
        //        var sentiment = await openAiClient.GetSentimentAsync(inputData.SentimentText);
        //        var predictedSentiment = MapSentiment(sentiment);
        //        var prediction = new SentimentPrediction
        //        {
        //            SentimentText = inputData.SentimentText,
        //            PredictedSentiment = predictedSentiment,
        //            Score = new float[] { } // You can add scores if needed
        //        };
        //        results.Add(prediction);
        //    }

        //    // Save results to Excel
        //    SaveResultsToExcel(results, outputFilePath);

        //}

    }

    public class OpenAIResponse
    {
        public string Sentiment { get; set; }
    }

    public class OpenAIClient
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;

        public OpenAIClient(string apiKey)
        {
            _httpClient = new HttpClient();
            _apiKey = apiKey;
        }

        public async Task<string> GetSentimentAsync(string text)
        {
            var requestBody = new
            {
                model = "text-davinci-003",
                prompt = $"Analyze the sentiment of the following text: \"{text}\". Respond with 'Positive', 'Neutral', or 'Negative'.",
                max_tokens = 10
            };

            var content = new StringContent(JsonConvert.SerializeObject(requestBody), Encoding.UTF8, "application/json");
            _httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _apiKey);

            try
            {
                var response = await _httpClient.PostAsync("https://api.openai.com/v1/completions", content);
                var responseBody = await response.Content.ReadAsStringAsync();
                // Log the response status code and body
                Console.WriteLine($"Response Status Code: {response.StatusCode}");
                Console.WriteLine($"Response Body: {responseBody}");

                response.EnsureSuccessStatusCode();

                var openAIResponse = JsonConvert.DeserializeObject<OpenAIResponse>(responseBody);

                return openAIResponse.Sentiment;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return "";
            }

        }
    }
}
