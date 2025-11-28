using System;

namespace OutlookAI
{
    public class CorrelationContext
    {
        public string CorrelationId { get; set; }
        public string OperationName { get; set; }
        public DateTime StartTime { get; set; }

        public CorrelationContext(string correlationId, string operationName = null)
        {
            CorrelationId = correlationId;
            OperationName = operationName;
            StartTime = DateTime.Now;
        }
    }
}
