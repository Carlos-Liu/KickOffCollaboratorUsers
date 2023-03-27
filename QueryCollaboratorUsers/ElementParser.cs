using System;
using System.Web;
using HtmlAgilityPack;

namespace QueryCollaboratorUsers
{
  internal static class ElementParser
  {
    public static string GetLogUri(HtmlNode node)
    {
      /*
       * <a class="gridHyperlink" href="go?step=AdminPreferences&page=Admin&userid=3"
       * title="Edit account information">[Edit]</a>
       * <a class="gridHyperlink"
       * href="go?component=ErrorsAndMessages&pageNumber=1&itemsPerPage=100&action=force-log-off&userListSort=LOGIN_ASC&step=AdminUsers&page=Admin&userListFilter=SHOW_ACTIVE_PER_LAST_HOUR&userid=3&ErrorsAndMessages_fingerPrint=861101"
       * title="Log this user off forceable and immediately">[Log Off]</a>
       */
      var logOffNode =node.ChildNodes[1];
      return logOffNode.Attributes["href"].Value;
    }

    public static DateTime GetLastActivityTime(string innerHtml)
    {
      // 2020-06-22 10&#x3a;14&#x3a;01
      // 2020-06-22 10&#x3a;17&#x3a;45
      var decodedHtml = HttpUtility.HtmlDecode(innerHtml);
      return DateTime.Parse(decodedHtml);
    }

    public static string GetEmail(HtmlNode node)
    {
      // mailto:xx@xx.com
      var html = node.ChildNodes[0].Attributes["href"].Value;
      return html.Split(':')[1];
    }
  }
}
