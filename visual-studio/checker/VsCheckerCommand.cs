using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace VsCheckerExtension
{
  internal sealed class VsCheckerCommand
  {
    // 匹配 using 行（支持 global using、static、alias）
    private static readonly Regex UsingRegex = new Regex(
      @"^\s*(global\s+)?using\s+([^;]+);",
      RegexOptions.Compiled);

    // 匹配 namespace 声明
    private static readonly Regex NamespaceRegex = new Regex(
      @"^\s*namespace\s+([A-Za-z_][A-Za-z0-9_.]*)",
      RegexOptions.Multiline | RegexOptions.Compiled);

    private static readonly Guid CommandSetGuid = new Guid("C09B25C8-8E90-4F0B-9C7E-7AA4379F3F9B");
    private const int CommandId = 0x0100;

    private readonly AsyncPackage package;
    private DTE2 dte;

    private VsCheckerCommand(AsyncPackage package, OleMenuCommandService commandService)
    {
      this.package = package;
      var menuCommandId = new CommandID(CommandSetGuid, CommandId);
      var menuItem = new MenuCommand(Execute, menuCommandId);
      commandService.AddCommand(menuItem);
    }

    public static async Task InitializeAsync(AsyncPackage package)
    {
      await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
      var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
      if (commandService != null)
      {
        new VsCheckerCommand(package, commandService);
      }
    }

    private void Execute(object sender, EventArgs e)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      EnsureDte();
      if (dte?.ActiveDocument == null)
      {
        return;
      }

      var document = dte.ActiveDocument;
      var ext = Path.GetExtension(document.FullName);
      if (!".cs".Equals(ext, StringComparison.OrdinalIgnoreCase))
      {
        return;
      }

      var textDocument = document.Object("TextDocument") as TextDocument;
      if (textDocument == null)
      {
        return;
      }

      var content = GetDocumentText(textDocument);
      var topNamespace = GetTopNamespace(content);
      var updated = SortUsings(content, topNamespace);
      updated = AppendTemplate(updated);

      if (!string.Equals(content, updated, StringComparison.Ordinal))
      {
        SetDocumentText(textDocument, updated);
        document.Save();
      }
    }

    private static string GetDocumentText(TextDocument textDocument)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      var start = textDocument.StartPoint.CreateEditPoint();
      return start.GetText(textDocument.EndPoint);
    }

    private static void SetDocumentText(TextDocument textDocument, string text)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      var start = textDocument.StartPoint.CreateEditPoint();
      start.ReplaceText(
        textDocument.EndPoint,
        text,
        (int)vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
    }

    // 从文件内容解析顶级 namespace 的第一段
    private static string GetTopNamespace(string content)
    {
      var match = NamespaceRegex.Match(content);
      if (!match.Success)
      {
        return string.Empty;
      }

      var full = match.Groups[1].Value.Trim();
      return full.Split('.')[0];
    }

    // 排序 using 声明
    private static string SortUsings(string content, string topNamespace)
    {
      var lineEnding = DetectLineEnding(content);
      var lines = content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

      var usingLines = new List<string>();
      var startIndex = -1;
      var endIndex = -1;
      var inBlockComment = false;

      for (var i = 0; i < lines.Length; i++)
      {
        var trimmed = lines[i].Trim();

        if (startIndex < 0)
        {
          if (IsIgnorableLine(trimmed, ref inBlockComment))
          {
            continue;
          }

          if (UsingRegex.IsMatch(trimmed))
          {
            startIndex = i;
            endIndex = i;
            usingLines.Add(trimmed);
            continue;
          }

          // 遇到非 using、非可忽略行，说明没有 using 区域
          return content;
        }

        // 已进入 using 区域
        if (UsingRegex.IsMatch(trimmed))
        {
          endIndex = i;
          usingLines.Add(trimmed);
          continue;
        }

        if (string.IsNullOrWhiteSpace(trimmed))
        {
          continue;
        }

        // 遇到非空非 using 行，using 区域结束
        break;
      }

      if (usingLines.Count == 0)
      {
        return content;
      }

      var sorted = usingLines
        .Select(line => new { Line = line, Group = GetGroup(line, topNamespace) })
        .OrderBy(x => x.Group)
        .ThenBy(x => x.Line.Length)
        .ThenBy(x => x.Line, StringComparer.Ordinal)
        .Select(x => x.Line)
        .ToList();

      var result = lines.Take(startIndex)
        .Concat(sorted)
        .Concat(lines.Skip(endIndex + 1));

      return string.Join(lineEnding, result);
    }

    // 计算 using 行所属的分组序号
    private static int GetGroup(string usingLine, string topNamespace)
    {
      var ns = ExtractNamespace(usingLine);
      if (string.IsNullOrEmpty(ns))
      {
        return 2;
      }

      var firstSegment = ns.Split('.')[0];

      if ("System".Equals(firstSegment, StringComparison.Ordinal))
      {
        return 0;
      }

      if ("Microsoft".Equals(firstSegment, StringComparison.Ordinal))
      {
        return 1;
      }

      if (!string.IsNullOrEmpty(topNamespace) &&
          topNamespace.Equals(firstSegment, StringComparison.Ordinal))
      {
        return 3;
      }

      return 2;
    }

    // 从 using 行提取 namespace 值
    private static string ExtractNamespace(string line)
    {
      var match = UsingRegex.Match(line);
      if (!match.Success)
      {
        return string.Empty;
      }

      var value = match.Groups[2].Value.Trim();

      // using static Foo.Bar;
      if (value.StartsWith("static ", StringComparison.Ordinal))
      {
        value = value.Substring(7).Trim();
      }

      // using global::Foo;
      if (value.StartsWith("global::", StringComparison.Ordinal))
      {
        value = value.Substring(8).Trim();
      }

      // using Alias = Foo.Bar;
      var eqIndex = value.IndexOf('=');
      if (eqIndex >= 0)
      {
        value = value.Substring(eqIndex + 1).Trim();
      }

      return value;
    }

    // 在文件末尾追加校验模板
    private string AppendTemplate(string content)
    {
      var options = (VsCheckerOptions)package.GetDialogPage(typeof(VsCheckerOptions));
      var template = options?.TemplateText ?? "//// Checked by XuRui @{now:o}";
      var nowText = DateTimeOffset.Now.ToString("o", CultureInfo.InvariantCulture);
      var line = template.Replace("{now:o}", nowText);
      var lineEnding = DetectLineEnding(content);

      if (!content.EndsWith("\n", StringComparison.Ordinal))
      {
        content += lineEnding;
      }

      return content + lineEnding + line + lineEnding;
    }

    private static bool IsIgnorableLine(string trimmed, ref bool inBlockComment)
    {
      if (inBlockComment)
      {
        if (trimmed.IndexOf("*/", StringComparison.Ordinal) >= 0)
        {
          inBlockComment = false;
        }

        return true;
      }

      if (trimmed.StartsWith("/*", StringComparison.Ordinal))
      {
        if (trimmed.IndexOf("*/", StringComparison.Ordinal) < 0)
        {
          inBlockComment = true;
        }

        return true;
      }

      if (trimmed.StartsWith("//", StringComparison.Ordinal))
      {
        return true;
      }

      if (trimmed.StartsWith("#", StringComparison.Ordinal))
      {
        return true;
      }

      if (trimmed.StartsWith("extern alias", StringComparison.Ordinal))
      {
        return true;
      }

      return trimmed.Length == 0;
    }

    private static string DetectLineEnding(string content)
    {
      return content.IndexOf("\r\n", StringComparison.Ordinal) >= 0 ? "\r\n" : "\n";
    }

    private void EnsureDte()
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      if (dte == null)
      {
        dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
      }
    }
  }
}
