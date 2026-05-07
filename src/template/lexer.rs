//! 模板词法分析器 (Lexer)
//!
//! 将模板字符串转换为词法单元(LexToken)流，供Parser使用。

/// 运算符类型
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Operator {
    Add,  // +
    Sub,  // -
    Mul,  // *
    Div,  // /
}

impl Operator {
    pub fn from_char(c: char) -> Option<Self> {
        match c {
            '+' => Some(Operator::Add),
            '-' => Some(Operator::Sub),
            '*' => Some(Operator::Mul),
            '/' => Some(Operator::Div),
            _ => None,
        }
    }

    pub fn to_char(&self) -> char {
        match self {
            Operator::Add => '+',
            Operator::Sub => '-',
            Operator::Mul => '*',
            Operator::Div => '/',
        }
    }
}

/// 词法单元类型
#[derive(Debug, Clone, PartialEq)]
pub enum LexToken {
    // 结构符号
    /// 标记开始 {{{
    MarkerOpen,
    /// 标记结束 }}}
    MarkerClose,
    /// 左括号 (
    ParenOpen,
    /// 右括号 )
    ParenClose,

    // 修饰符
    /// 行迭代 @
    AtMark,
    /// 变量运算 #
    HashMark,
    /// Excel公式 =
    EqualSign,
    /// 条件判断 ? (预留)
    QuestionMark,

    // 标识符
    /// 标识符 (tag, field, func名)
    Identifier(String),

    // 分隔符
    /// 点号 .
    Dot,
    /// 冒号 :
    Colon,
    /// 逗号 ,
    Comma,

    // 运算符
    /// 运算符 +, -, *, /
    Operator(Operator),

    // 字面量
    /// 数字字面量
    Number(f64),
    /// 字符串字面量 (预留)
    StringLit(String),

    // 格式码
    /// 格式码内容 (如 float:2, pct, int)
    FormatSpec(String),

    // 其他
    /// 普通文本 (非标记内容)
    Text(String),
    /// 结束标记
    EOF,
}

impl LexToken {
    /// 判断是否是修饰符
    pub fn is_modifier(&self) -> bool {
        matches!(self, LexToken::AtMark | LexToken::HashMark | LexToken::EqualSign | LexToken::QuestionMark)
    }

    /// 判断是否是运算符
    pub fn is_operator(&self) -> bool {
        matches!(self, LexToken::Operator(_))
    }

    /// 判断是否是加减运算符
    pub fn is_add_sub(&self) -> bool {
        matches!(self, LexToken::Operator(Operator::Add) | LexToken::Operator(Operator::Sub))
    }

    /// 判断是否是乘除运算符
    pub fn is_mul_div(&self) -> bool {
        matches!(self, LexToken::Operator(Operator::Mul) | LexToken::Operator(Operator::Div))
    }
}

/// 词法分析器
pub struct Lexer;

impl Lexer {
    /// 将输入字符串转换为词法单元流
    pub fn tokenize(input: &str) -> Vec<LexToken> {
        let mut tokens = Vec::new();
        let mut pos = 0;
        let chars: Vec<char> = input.chars().collect();

        while pos < chars.len() {
            // 检查是否是标记开始 {{{
            if pos + 1 < chars.len() && chars[pos] == '{' && chars[pos + 1] == '{' {
                // 先收集之前的文本（如果有）
                // 由于我们是在标记开始处，前面的文本应该在上一轮已经收集

                // 进入标记内部解析
                pos += 2; // 跳过 {{
                tokens.push(LexToken::MarkerOpen);

                // 解析标记内部内容
                Self::parse_marker_content(&chars, &mut pos, &mut tokens);
            } else {
                // 收集普通文本直到下一个标记开始
                let text_start = pos;
                while pos < chars.len() {
                    if pos + 1 < chars.len() && chars[pos] == '{' && chars[pos + 1] == '{' {
                        break;
                    }
                    pos += 1;
                }
                if pos > text_start {
                    let text: String = chars[text_start..pos].iter().collect();
                    tokens.push(LexToken::Text(text));
                }
            }
        }

        tokens.push(LexToken::EOF);
        tokens
    }

    /// 解析标记内部内容 {{ ... }}}
    fn parse_marker_content(chars: &[char], pos: &mut usize, tokens: &mut Vec<LexToken>) {
        // 解析修饰符
        while *pos < chars.len() {
            let c = chars[*pos];

            // 检查标记结束 }}}
            if *pos + 1 < chars.len() && c == '}' && chars[*pos + 1] == '}' {
                *pos += 2; // 跳过 }}
                tokens.push(LexToken::MarkerClose);
                return;
            }

            // 修饰符
            match c {
                '@' => {
                    tokens.push(LexToken::AtMark);
                    *pos += 1;
                }
                '#' => {
                    tokens.push(LexToken::HashMark);
                    *pos += 1;
                }
                '=' => {
                    tokens.push(LexToken::EqualSign);
                    *pos += 1;
                }
                '?' => {
                    tokens.push(LexToken::QuestionMark);
                    *pos += 1;
                }
                _ => break,
            }
        }

        // 解析表达式部分
        Self::parse_expression(chars, pos, tokens);

        // 解析格式码 :format
        if *pos < chars.len() && chars[*pos] == ':' {
            *pos += 1; // 跳过 :
            tokens.push(LexToken::Colon);
            Self::parse_format_spec(chars, pos, tokens);
        }

        // 期望标记结束 }}}
        if *pos + 1 < chars.len() && chars[*pos] == '}' && chars[*pos + 1] == '}' {
            *pos += 2;
            tokens.push(LexToken::MarkerClose);
        }
        // 如果没有正确结束，暂时忽略错误
    }

    /// 解析表达式部分
    fn parse_expression(chars: &[char], pos: &mut usize, tokens: &mut Vec<LexToken>) {
        while *pos < chars.len() {
            let c = chars[*pos];

            // 检查标记结束或格式码开始
            if *pos + 1 < chars.len() && c == '}' && chars[*pos + 1] == '}' {
                return; // 标记结束
            }
            if c == ':' {
                return; // 格式码开始
            }

            // 左括号
            if c == '(' {
                tokens.push(LexToken::ParenOpen);
                *pos += 1;
                continue;
            }

            // 右括号
            if c == ')' {
                tokens.push(LexToken::ParenClose);
                *pos += 1;
                continue;
            }

            // 点号
            if c == '.' {
                tokens.push(LexToken::Dot);
                *pos += 1;
                continue;
            }

            // 逗号
            if c == ',' {
                tokens.push(LexToken::Comma);
                *pos += 1;
                continue;
            }

            // 运算符
            if let Some(op) = Operator::from_char(c) {
                // 注意：负数的 - 需要特殊处理
                // 如果 - 前面是数字、标识符或 )，则是运算符
                // 如果 - 前面是运算符或 ( 或开头，则是负号（合并到数字）
                tokens.push(LexToken::Operator(op));
                *pos += 1;
                continue;
            }

            // 数字 (包括负数)
            if c.is_ascii_digit() || (c == '-' && *pos + 1 < chars.len() && chars[*pos + 1].is_ascii_digit()) {
                let num = Self::parse_number(chars, pos);
                tokens.push(LexToken::Number(num));
                continue;
            }

            // 标识符 (字母开头，后续可以是字母、数字、下划线)
            if c.is_alphabetic() || c == '_' {
                let ident = Self::parse_identifier(chars, pos);
                tokens.push(LexToken::Identifier(ident));
                continue;
            }

            // 花括号占位符 {r}, {r+1}, {r-1} (Excel公式中的行号)
            if c == '{' && *pos + 1 < chars.len() && chars[*pos + 1] == 'r' {
                // 解析 {r...} 占位符，作为标识符处理
                let start = *pos;
                *pos += 1; // 跳过 {
                while *pos < chars.len() && chars[*pos] != '}' {
                    *pos += 1;
                }
                if *pos < chars.len() {
                    *pos += 1; // 跳过 }
                }
                let placeholder: String = chars[start..*pos].iter().collect();
                tokens.push(LexToken::Identifier(placeholder));
                continue;
            }

            // 空格忽略
            if c == ' ' || c == '\t' {
                *pos += 1;
                continue;
            }

            // 其他字符，暂时跳过
            *pos += 1;
        }
    }

    /// 解析数字
    fn parse_number(chars: &[char], pos: &mut usize) -> f64 {
        let start = *pos;
        let mut has_dot = false;

        // 处理负号
        if chars[*pos] == '-' {
            *pos += 1;
        }

        while *pos < chars.len() {
            let c = chars[*pos];
            if c.is_ascii_digit() {
                *pos += 1;
            } else if c == '.' && !has_dot {
                has_dot = true;
                *pos += 1;
            } else {
                break;
            }
        }

        let num_str: String = chars[start..*pos].iter().collect();
        num_str.parse::<f64>().unwrap_or(0.0)
    }

    /// 解析标识符
    fn parse_identifier(chars: &[char], pos: &mut usize) -> String {
        let start = *pos;

        while *pos < chars.len() {
            let c = chars[*pos];
            if c.is_alphanumeric() || c == '_' {
                *pos += 1;
            } else {
                break;
            }
        }

        chars[start..*pos].iter().collect()
    }

    /// 解析格式码
    fn parse_format_spec(chars: &[char], pos: &mut usize, tokens: &mut Vec<LexToken>) {
        let start = *pos;

        while *pos < chars.len() {
            let c = chars[*pos];

            // 检查标记结束 }}}
            if *pos + 1 < chars.len() && c == '}' && chars[*pos + 1] == '}' {
                break;
            }

            // 格式码允许的字符：字母、数字、冒号、括号、符号
            if c.is_alphanumeric() || c == ':' || c == '(' || c == ')' || c == '$' || c == '¥' || c == '-' || c == '_' {
                *pos += 1;
            } else {
                break;
            }
        }

        if *pos > start {
            let format_str: String = chars[start..*pos].iter().collect();
            tokens.push(LexToken::FormatSpec(format_str));
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tokenize_simple_variable() {
        let tokens = Lexer::tokenize("{{tag.field}}");
        // MarkerOpen, Identifier(tag), Dot, Identifier(field), MarkerClose, EOF
        assert_eq!(tokens.len(), 6);
        assert!(matches!(tokens[0], LexToken::MarkerOpen));
        if let LexToken::Identifier(s) = &tokens[1] {
            assert_eq!(s, "tag");
        }
        assert!(matches!(tokens[2], LexToken::Dot));
        if let LexToken::Identifier(s) = &tokens[3] {
            assert_eq!(s, "field");
        }
        assert!(matches!(tokens[4], LexToken::MarkerClose));
        assert!(matches!(tokens[5], LexToken::EOF));
    }

    #[test]
    fn test_tokenize_row_iterate() {
        let tokens = Lexer::tokenize("{{@tag.field:float:2}}");
        assert!(matches!(tokens[0], LexToken::MarkerOpen));
        assert!(matches!(tokens[1], LexToken::AtMark));
        assert!(matches!(tokens[2], LexToken::Identifier(_)));
    }

    #[test]
    fn test_tokenize_formula() {
        let tokens = Lexer::tokenize("{{=A{r}*B{r}}}");
        assert!(matches!(tokens[0], LexToken::MarkerOpen));
        assert!(matches!(tokens[1], LexToken::EqualSign));
        // Formula content is parsed - A{r} becomes identifier "A{r}"
        // Check that there are identifier tokens and a multiplication operator
        let has_ident = tokens.iter().any(|t| matches!(t, LexToken::Identifier(_)));
        assert!(has_ident);
        let has_mul_op = tokens.iter().any(|t| matches!(t, LexToken::Operator(Operator::Mul)));
        assert!(has_mul_op);
    }

    #[test]
    fn test_tokenize_aggregate() {
        let tokens = Lexer::tokenize("{{sum(tag.field)}}");
        assert!(matches!(tokens[0], LexToken::MarkerOpen));
        assert!(matches!(tokens[1], LexToken::Identifier(_))); // sum
        assert!(matches!(tokens[2], LexToken::ParenOpen));
    }

    #[test]
    fn test_tokenize_mixed_text() {
        let tokens = Lexer::tokenize("Hello {{name}} World");
        // Text("Hello "), MarkerOpen, Identifier("name"), MarkerClose, Text(" World"), EOF
        if let LexToken::Text(s) = &tokens[0] {
            assert_eq!(s.trim(), "Hello");
        } else {
            panic!("Expected Text at position 0, got {:?}", tokens[0]);
        }
        assert!(matches!(tokens[1], LexToken::MarkerOpen));
        if let LexToken::Identifier(s) = &tokens[2] {
            assert_eq!(s, "name");
        }
        assert!(matches!(tokens[3], LexToken::MarkerClose));
        if let LexToken::Text(s) = &tokens[4] {
            assert_eq!(s.trim(), "World");
        }
    }

    #[test]
    fn test_tokenize_binop() {
        let tokens = Lexer::tokenize("{{#(tag1.f1*tag2.f2)}}");
        assert!(matches!(tokens[0], LexToken::MarkerOpen));
        assert!(matches!(tokens[1], LexToken::HashMark));
        assert!(matches!(tokens[2], LexToken::ParenOpen));
        assert!(matches!(tokens[3], LexToken::Identifier(_))); // tag1
        assert!(matches!(tokens[4], LexToken::Dot));
        assert!(matches!(tokens[5], LexToken::Identifier(_))); // f1
        assert!(matches!(tokens[6], LexToken::Operator(Operator::Mul)));
    }
}