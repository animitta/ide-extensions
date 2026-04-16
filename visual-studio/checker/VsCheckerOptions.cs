using System;
using System.ComponentModel;
using Microsoft.VisualStudio.Shell;

namespace VsCheckerExtension
{
  public class VsCheckerOptions : DialogPage
  {
    [Category("VsChecker")]
    [DisplayName("校验模板")]
    [Description("文件末尾插入的校验模板。{now:o} 会被替换为当前时间。")]
    public string TemplateText { get; set; } = "//// Checked by XuRui @{now:o}";
  }
}
