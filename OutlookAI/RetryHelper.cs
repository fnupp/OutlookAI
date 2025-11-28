using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace OutlookAI
{
    /// <summary>
    /// Helper class for retry logic with exponential backoff
    /// </summary>
    public static class RetryHelper
    {
        /// <summary>
        /// Executes an async function with retry logic and exponential backoff
        /// </summary>
        /// <typeparam name="T">Return type of the function</typeparam>
        /// <param name="func">The function to execute</param>
        /// <param name="maxAttempts">Maximum number of retry attempts</param>
        /// <param name="initialDelayMs">Initial delay in milliseconds</param>
        /// <param name="operationName">Name of the operation for logging</param>
        /// <returns>Result of the function</returns>
        public static async Task<T> ExecuteWithRetryAsync<T>(
            Func<Task<T>> func,
            int maxAttempts = 3,
            int initialDelayMs = 1000,
            string operationName = "Operation")
        {
            int attempt = 0;
            int delayMs = initialDelayMs;

            while (true)
            {
                attempt++;

                try
                {
                    ErrorLogger.LogInfo($"{operationName}: Attempt {attempt}/{maxAttempts}");
                    return await func().ConfigureAwait(false);
                }
                catch (Exception ex) when (IsRetriableException(ex) && attempt < maxAttempts)
                {
                    string errorType = GetErrorType(ex);
                    ErrorLogger.LogWarning(
                        $"{operationName}: {errorType} on attempt {attempt}/{maxAttempts}. " +
                        $"Retrying in {delayMs}ms...");

                    await Task.Delay(delayMs).ConfigureAwait(false);

                    // Exponential backoff with jitter
                    delayMs = CalculateNextDelay(delayMs, attempt);
                }
                catch (Exception ex)
                {
                    // Non-retriable exception or max attempts reached
                    string errorType = GetErrorType(ex);
                    ErrorLogger.LogError(
                        $"{operationName}: Failed after {attempt} attempt(s). Error: {errorType}",
                        ex);

                    throw new LLMCommunicationException(
                        $"{operationName} failed after {attempt} attempt(s): {errorType}",
                        ex);
                }
            }
        }

        /// <summary>
        /// Determines if an exception is retriable
        /// </summary>
        private static bool IsRetriableException(Exception ex)
        {
            // Timeout exceptions are retriable
            if (ex is TaskCanceledException || ex is TimeoutException)
                return true;

            // HTTP exceptions
            if (ex is HttpRequestException httpEx)
            {
                // Network errors are retriable
                if (httpEx.InnerException is WebException webEx)
                {
                    switch (webEx.Status)
                    {
                        case WebExceptionStatus.Timeout:
                        case WebExceptionStatus.ConnectFailure:
                        case WebExceptionStatus.NameResolutionFailure:
                        case WebExceptionStatus.ProxyNameResolutionFailure:
                        case WebExceptionStatus.SendFailure:
                        case WebExceptionStatus.ReceiveFailure:
                        case WebExceptionStatus.ConnectionClosed:
                        case WebExceptionStatus.KeepAliveFailure:
                        case WebExceptionStatus.PipelineFailure:
                            return true;
                    }
                }

                // Specific HTTP status codes that are retriable
                if (httpEx.Message.Contains("503") || // Service Unavailable
                    httpEx.Message.Contains("502") || // Bad Gateway
                    httpEx.Message.Contains("504") || // Gateway Timeout
                    httpEx.Message.Contains("429"))   // Too Many Requests
                {
                    return true;
                }

                return true; // Retry other HTTP exceptions
            }

            // Generic exceptions that might be network related
            if (ex.Message.Contains("Unable to connect") ||
                ex.Message.Contains("No connection") ||
                ex.Message.Contains("network") ||
                ex.Message.Contains("timeout"))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Gets a friendly error type description
        /// </summary>
        private static string GetErrorType(Exception ex)
        {
            if (ex is TaskCanceledException)
                return "Timeout";

            if (ex is TimeoutException)
                return "Request timeout";

            if (ex is HttpRequestException httpEx)
            {
                if (httpEx.InnerException is WebException webEx)
                {
                    return $"Network error ({webEx.Status})";
                }

                if (httpEx.Message.Contains("503"))
                    return "Service unavailable";
                if (httpEx.Message.Contains("502"))
                    return "Bad gateway";
                if (httpEx.Message.Contains("504"))
                    return "Gateway timeout";
                if (httpEx.Message.Contains("429"))
                    return "Rate limit exceeded";

                return "HTTP request failed";
            }

            return ex.GetType().Name;
        }

        /// <summary>
        /// Calculates the next delay with exponential backoff and jitter
        /// </summary>
        private static int CalculateNextDelay(int currentDelayMs, int attempt)
        {
            // Exponential backoff: double the delay each time
            int nextDelay = currentDelayMs * 2;

            // Add jitter (random variation of ±20%)
            Random random = new Random();
            double jitterFactor = 0.8 + (random.NextDouble() * 0.4); // 0.8 to 1.2
            nextDelay = (int)(nextDelay * jitterFactor);

            // Cap at 30 seconds
            return Math.Min(nextDelay, 30000);
        }
    }

    /// <summary>
    /// Custom exception for LLM communication failures
    /// </summary>
    public class LLMCommunicationException : Exception
    {
        public LLMCommunicationException(string message) : base(message) { }
        public LLMCommunicationException(string message, Exception innerException) : base(message, innerException) { }
    }
}
