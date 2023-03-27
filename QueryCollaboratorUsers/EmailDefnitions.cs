using System.Collections.Generic;

namespace QueryCollaboratorUsers
{

  public class EmailEntity
  {
    public string SenderEmailAddress { get; set; }
    public string SenderName { get; set; }
    public string ToEmailAddress { get; set; }
    public string ToName { get; set; }
    public string Title { get; set; }
    public string EmailTextContent { get; set; }
    public string EmailHtmlContent { get; set; }
  }
  public class Email
  {
    public string SenderEmailAddress { get; set; }
    public string SenderName { get; set; }
    public List<ToEmailFullAddressName> ToEmailFullAddressNames { get; set; }
    public string Title { get; set; }
    public string EmailTextContent { get; set; }
    public string EmailHtmlContent { get; set; }
  }

  public class ToEmailFullAddressName
  {
    public ToEmailFullAddressName(string toEmailAddress, string toName)
    {
      ToEmailAddress = toEmailAddress;
      ToName = toName;
    }

    public string ToEmailAddress { get; set; }
    public string ToName { get; set; }
  }
}
