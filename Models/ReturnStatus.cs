namespace EFormApi.Models
{
    public enum ReturnStatus
    {
        OK,
        NG,
    }

    public enum ReturnDetail
    {
        Sucessfull,
        TimeOut,
        WrongSendData

    }

    public class ReturnModel
    {
        public ReturnStatus Status { get; set; } 
        public ReturnDetail ReturnDetail { get; set; }
        public string Target { get; set; } = null!;
    }
}
