using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public class CodeHighlighter
    {
        // ... 其他代码保持不变 ...

        private void InitializePatterns()
        {
            languagePatterns = new Dictionary<string, List<(string pattern, RegexOptions options, string type)>>
            {
                // ... 原有语言模式保持不变 ...
                
                // 添加C++语言支持
                {"cpp", new List<(string, RegexOptions, string)>
                    {
                        // 字符串（支持宽字符串和转义字符）
                        (@"L?""([^""\n\\]|\\.)*""", RegexOptions.None, "string"),
                        // 数字（支持各种格式）
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?[fFlLuU]*\b", RegexOptions.None, "number"),
                        // 注释（支持单行和多行）
                        (@"/\*[\s\S]*?\*/", RegexOptions.None, "comment"),
                        (@"//[^\n]*", RegexOptions.None, "comment"),
                        // 预处理器指令
                        (@"#\s*(include|define|ifdef|ifndef|if|else|elif|endif|pragma|undef|error|line|using)\b", RegexOptions.None, "keyword"),
                        // 关键字（C++11/14/17特性）
                        (@"\b(alignas|alignof|and|and_eq|asm|auto|bitand|bitor|bool|break|case|catch|char|char8_t|char16_t|char32_t|class|compl|concept|const|consteval|constexpr|const_cast|continue|co_await|co_return|co_yield|decltype|default|delete|do|double|dynamic_cast|else|enum|explicit|export|extern|false|float|for|friend|goto|if|inline|int|long|mutable|namespace|new|noexcept|not|not_eq|nullptr|operator|or|or_eq|private|protected|public|register|reinterpret_cast|requires|return|short|signed|sizeof|static|static_assert|static_cast|struct|switch|template|this|thread_local|throw|true|try|typedef|typeid|typename|union|unsigned|using|virtual|void|volatile|wchar_t|while|xor|xor_eq)\b", RegexOptions.None, "keyword"),
                    }
                },
                // 添加C语言支持
                {"c", new List<(string, RegexOptions, string)>
                    {
                        // 字符串（支持转义字符）
                        (@"L?""([^""\n\\]|\\.)*""", RegexOptions.None, "string"),
                        // 数字（支持各种格式）
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?[fFlLuU]*\b", RegexOptions.None, "number"),
                        // 注释（支持单行和多行）
                        (@"/\*[\s\S]*?\*/", RegexOptions.None, "comment"),
                        (@"//[^\n]*", RegexOptions.None, "comment"),
                        // 预处理器指令
                        (@"#\s*(include|define|ifdef|ifndef|if|else|elif|endif|pragma|undef|error|line)\b", RegexOptions.None, "keyword"),
                        // 关键字（ANSI C标准）
                        (@"\b(auto|break|case|char|const|continue|default|do|double|else|enum|extern|float|for|goto|if|inline|int|long|register|return|short|signed|sizeof|static|struct|switch|typedef|union|unsigned|void|volatile|while|_Alignas|_Alignof|_Atomic|_Bool|_Complex|_Generic|_Imaginary|_Noreturn|_Static_assert|_Thread_local)\b", RegexOptions.None, "keyword"),
                    }
                }
            };
        }

        private void InitializeAliases()
        {
            languageAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // ... 原有别名保持不变 ...
                
                // 添加C++别名
                {"c++", "cpp"},
                {"cxx", "cpp"},
                {"cc", "cpp"},
                {"hpp", "cpp"},
                {"hxx", "cpp"},
                {"hh", "cpp"},
                
                // 添加C语言别名
                {"c", "c"},
                {"h", "c"}
            };
        }

        // ... 其他代码保持不变 ...
    }
}
