// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Agents.Core.Models;
using Microsoft.Agents.CopilotStudio.Client;
using System.Text.Json.Nodes;
using System.Text.Json;
using Newtonsoft.Json;
using ClosedXML.Excel;
using JsonSerializer = System.Text.Json.JsonSerializer;
using System.ComponentModel;

namespace CopilotStudioClientSample;

internal class ChatConsoleService(CopilotClient copilotClient) : IHostedService
{
    private static readonly string ExcelFileName = $"Response_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx";

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        bool runInBatch = true;
        Console.WriteLine("\nChoose an option:");
        Console.WriteLine("1. Ask your own questions");
        Console.WriteLine("2. Run batch from questions.xlsx");
        Console.Write("\nEnter your choice (defaulting to batch in 15 seconds): ");

        string? input = null;
        Task inputTask = Task.Run(() => input = Console.ReadLine());
        bool completedInTime = await Task.WhenAny(inputTask, Task.Delay(15000)) == inputTask;

        if (completedInTime)
        {
            if (input?.Trim() == "1")
            {
                runInBatch = false;
            }
        }
        
        if (runInBatch)
        {
            Console.WriteLine("\nRunning batch mode...");
            await RunBatchMode(cancellationToken);
        }

        else
        {
            Console.WriteLine("\nRunning Interactive mode mode...");
            await RunInteractiveMode(cancellationToken);
        }
        Console.WriteLine("\nExecution completed. Press Enter to exit.");
    Console.ReadLine();
    }

    private async Task RunInteractiveMode(CancellationToken cancellationToken)
    {
        System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
        Console.Write("\nUser> ");

        await foreach (Activity act in copilotClient.StartConversationAsync(emitStartConversationEvent: true, cancellationToken: cancellationToken))
        {
            System.Diagnostics.Trace.WriteLine($">>>>MessageLoop Duration: {sw.Elapsed.ToDurationString()}");
            sw.Restart();
            if (act is null) throw new InvalidOperationException("Activity is null");
            Console.WriteLine("\nAgent> " + act.Text);
        }

        while (!cancellationToken.IsCancellationRequested)
        {
            Console.Write("\nUser> ");
            string question = Console.ReadLine()!;
            Console.WriteLine("\nAgent>");
            sw.Restart();
            await foreach (Activity act in copilotClient.AskQuestionAsync(question, null, cancellationToken))
            {
                System.Diagnostics.Trace.WriteLine($">>>>MessageLoop Duration: {sw.Elapsed.ToDurationString()}");
                await PrintActivityAsync(act, cancellationToken);
                sw.Restart();
            }
        }

        sw.Stop();
    }

    private async Task RunBatchMode(CancellationToken cancellationToken)
    {
        const string ExcelInputFile = "questions.xlsx";
        const string SheetName = "Questions";
        const int QuestionColumnIndex = 1; // Column A
        const int StartRow = 1; // Assuming row 1 is header

        if (!File.Exists(ExcelInputFile))
        {
            Console.WriteLine($"Error: {ExcelInputFile} not found.");
            return;
        }

        var questions = new List<string>();

        using (var workbook = new XLWorkbook(ExcelInputFile))
        {
            var worksheet = workbook.Worksheet(SheetName);
            var row = StartRow;
            while (!string.IsNullOrWhiteSpace(worksheet.Cell(row, QuestionColumnIndex).GetString()))
            {
                questions.Add(worksheet.Cell(row, QuestionColumnIndex).GetString());
                row++;
            }
        }

        if (questions.Count == 0)
        {
            Console.WriteLine("No questions found in the Excel file.");
            return;
        }

        var outputWorkbook = new XLWorkbook();
        var outputSheet = outputWorkbook.Worksheets.Add("Results");
        outputSheet.Cell(1, 1).Value = "Question";
        outputSheet.Cell(1, 2).Value = "Response";
        outputSheet.Cell(1, 3).Value = "Conversation id";
        outputSheet.Cell(1, 4).Value = "Timestamp";
        outputSheet.Cell(1, 5).Value = "Response Log";

        int outputRow = 2;

        await foreach (Activity act in copilotClient.StartConversationAsync(true, cancellationToken))
        {
            Console.WriteLine("Agent> " + act.Text);
            outputSheet.Cell(outputRow, 1).Value = "System Start";
            outputSheet.Cell(outputRow, 2).Value = act.Text;
            outputSheet.Cell(outputRow, 3).Value = act.Text;
            outputSheet.Cell(outputRow, 4).Value = act.Text;
            outputSheet.Cell(outputRow, 5).Value = DateTime.Now;
            outputRow++;
            break;
        }

        for (int i = 0; i < questions.Count; i++)
        {
            var question = questions[i];
            Console.WriteLine($"\nAsking question {i + 1} of {questions.Count}");
            Console.WriteLine($"User> {question}");
            string response = "";
            string responseLog = "";
            string conversationId = "";

            await foreach (Activity act in copilotClient.AskQuestionAsync(question, null, cancellationToken))
            {
                if (!string.IsNullOrEmpty(act.Text))
                {
                    Console.WriteLine("Agent> " + act.Text);
                    response += act.Text + "\n";
                }
               responseLog += JsonConvert.SerializeObject(act, Formatting.Indented) + "\n";
                if (act.Conversation != null)
                {
                    conversationId += act.Conversation.Id;
                }

            }

            string trimmedResponse = response.Length > 32767 ? response.Substring(0, 32767) : response;
            string trimmedResponseLog = responseLog.Length > 32767 ? responseLog.Substring(0, 32767) : responseLog;

            outputSheet.Cell(outputRow, 1).Value = question;
            outputSheet.Cell(outputRow, 2).Value = trimmedResponse;
            outputSheet.Cell(outputRow, 3).Value = conversationId;
            outputSheet.Cell(outputRow, 4).Value = DateTime.Now;
            outputSheet.Cell(outputRow, 5).Value = trimmedResponseLog;
            outputRow++;
        }

        outputWorkbook.SaveAs(ExcelFileName);
        Console.WriteLine($"\nResults saved to {ExcelFileName}");
    }

    private async Task PrintActivityAsync(IActivity act, CancellationToken cancellationToken)
    {
        switch (act.Type)
        {
            case "message":
                Console.WriteLine(act.Text);

                if (act.SuggestedActions?.Actions?.Count > 0)
                {
                    Console.WriteLine("Suggested actions:\n");
                    foreach (var action in act.SuggestedActions.Actions)
                        Console.WriteLine($"\t{action.Text}");
                }

                if (act.Attachments?.Count > 0)
                {
                    foreach (var attachment in act.Attachments)
                    {
                        if (attachment.ContentType == "application/vnd.microsoft.card.adaptive")
                        {
                            var userInput = HandleAdaptiveCard(attachment.Content);
                            if (userInput.Count > 0)
                            {
                                Console.WriteLine("\nSending your inputs to the agent...\n");
                                var userActivity = new Activity
                                {
                                    Type = "message",
                                    Text = JsonSerializer.Serialize(userInput),
                                    From = new ChannelAccount { Id = "user", Name = "User" },
                                    Recipient = new ChannelAccount { Id = "bot", Name = "Bot" },
                                    Conversation = act.Conversation,
                                    ReplyToId = act.Id
                                };

                                var userInputJson = JsonSerializer.Serialize(userInput);
                                await foreach (var followUp in copilotClient.AskQuestionAsync(userInputJson, null, cancellationToken))
                                {
                                    await PrintActivityAsync(followUp, cancellationToken);
                                }
                            }
                        }
                    }
                }
                break;

            case "typing":
                Console.Write(".");
                break;

            case "event":
                Console.Write("+");
                break;

            default:
                Console.Write($"[{act.Type}]");
                break;
        }
    }

    private static Dictionary<string, object> HandleAdaptiveCard(object content)
    {
        var inputs = new Dictionary<string, object>();
        var cardJson = content as JsonObject ?? JsonNode.Parse(content?.ToString()) as JsonObject;

        if (cardJson?["body"] is not JsonArray body)
        {
            Console.WriteLine("[!] Adaptive Card body is missing or malformed.");
            return inputs;
        }

        foreach (var item in body.OfType<JsonObject>())
        {
            var type = item["type"]?.ToString();
            var id = item["id"]?.ToString();
            if (string.IsNullOrWhiteSpace(id)) continue;

            var label = item["label"]?.ToString() ?? item["placeholder"]?.ToString() ?? id;

            switch (type)
            {
                case "Input.Text":
                    Console.Write($"{label}: ");
                    inputs[id] = Console.ReadLine();
                    break;

                case "Input.Number":
                    Console.Write($"{label} (number): ");
                    if (int.TryParse(Console.ReadLine(), out int numberValue))
                        inputs[id] = numberValue;
                    break;

                case "Input.ChoiceSet":
                    var choices = item["choices"]?.AsArray();
                    if (choices != null && choices.Count > 0)
                    {
                        Console.WriteLine($"{label}:");
                        for (int i = 0; i < choices.Count; i++)
                        {
                            Console.WriteLine($"  {i + 1}. {choices[i]?["title"]?.ToString()}");
                        }

                        Console.Write("Select option number: ");
                        if (int.TryParse(Console.ReadLine(), out int selectedIndex) &&
                            selectedIndex >= 1 && selectedIndex <= choices.Count)
                        {
                            inputs[id] = choices[selectedIndex - 1]?["value"]?.ToString();
                        }
                    }
                    break;

                case "Input.Toggle":
                    Console.Write($"{label} (yes/no): ");
                    string? toggleInput = Console.ReadLine()?.Trim().ToLower();
                    inputs[id] = (toggleInput == "yes" || toggleInput == "y")
                        ? item["valueOn"]?.ToString() ?? "true"
                        : item["valueOff"]?.ToString() ?? "false";
                    break;

                case "Input.Date":
                    Console.Write($"{label} (yyyy-MM-dd): ");
                    inputs[id] = Console.ReadLine();
                    break;

                case "Input.Time":
                    Console.Write($"{label} (HH:mm): ");
                    inputs[id] = Console.ReadLine();
                    break;
            }
        }

        return inputs;
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        System.Diagnostics.Trace.TraceInformation("Stopping");
        return Task.CompletedTask;
    }
}
