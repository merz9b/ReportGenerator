//! 语法分析器 (Parser)
//!
//! 基于词法单元流构建AST，支持运算符优先级和嵌套表达式。

use crate::template::ast::*;
use crate::template::lexer::{LexToken, Lexer, Operator};

/// 语法分析器
pub struct Parser {
    tokens: Vec<LexToken>,
    pos: usize,
}

impl Parser {
    /// 解析单元格内容，返回解析结果
    pub fn parse_cell(input: &str) -> ParsedCell {
        let tokens = Lexer::tokenize(input);
        let parser = Parser::new(tokens);
        let nodes = parser.parse_nodes();
        ParsedCell::from_nodes(nodes)
    }

    fn new(tokens: Vec<LexToken>) -> Self {
        Self { tokens, pos: 0 }
    }

    /// 解析所有节点
    fn parse_nodes(mut self) -> Vec<AstNode> {
        let mut nodes = Vec::new();

        while !self.is_at_end() {
            if let Some(node) = self.parse_node() {
                nodes.push(node);
            } else {
                self.advance();
            }
        }

        nodes
    }

    /// 解析单个节点
    fn parse_node(&mut self) -> Option<AstNode> {
        let token = self.peek().clone();
        match token {
            LexToken::MarkerOpen => Some(self.parse_marker()),
            LexToken::Text(s) => {
                self.advance();
                Some(AstNode::Text(s))
            },
            _ => None,
        }
    }

    /// 解析标记节点 {{ modifier? expr :format? }}}
    fn parse_marker(&mut self) -> AstNode {
        self.advance(); // 跳过 MarkerOpen

        // 解析修饰符
        let modifier = self.parse_modifier();

        // 对于公式修饰符，直接收集公式内容
        let expr = if matches!(modifier, Modifier::Formula | Modifier::RowFormula) {
            self.parse_excel_formula()
        } else {
            self.parse_expression()
        };

        // 解析格式码
        let format = self.parse_format();

        // 期望 MarkerClose
        if matches!(self.peek(), LexToken::MarkerClose) {
            self.advance();
        }

        AstNode::Marker { modifier, expr, format }
    }

    /// 解析 Excel 公式内容（收集直到 Colon 或 MarkerClose）
    fn parse_excel_formula(&mut self) -> Expression {
        let mut formula_parts: Vec<String> = Vec::new();

        while !self.is_at_end() {
            let token = self.peek().clone();

            // 遇到冒号（格式码开始）或标记结束，停止
            if matches!(token, LexToken::Colon | LexToken::MarkerClose) {
                break;
            }

            // 收集 token 内容
            match &token {
                LexToken::Identifier(s) => formula_parts.push(s.clone()),
                LexToken::Operator(op) => formula_parts.push(op.to_char().to_string()),
                LexToken::Number(n) => formula_parts.push(n.to_string()),
                LexToken::ParenOpen => formula_parts.push("(".to_string()),
                LexToken::ParenClose => formula_parts.push(")".to_string()),
                LexToken::Dot => formula_parts.push(".".to_string()),
                LexToken::Comma => formula_parts.push(".".to_string()),
                _ => {}
            }

            self.advance();
        }

        Expression::ExcelFormula { formula: formula_parts.join("") }
    }

    /// 解析修饰符
    fn parse_modifier(&mut self) -> Modifier {
        // 处理组合修饰符 @=
        if matches!(self.peek(), LexToken::AtMark) {
            self.advance();
            if matches!(self.peek(), LexToken::EqualSign) {
                self.advance();
                return Modifier::RowFormula;
            }
            return Modifier::RowIterate;
        }

        match self.peek() {
            LexToken::HashMark => {
                self.advance();
                Modifier::VariableOp
            },
            LexToken::EqualSign => {
                self.advance();
                Modifier::Formula
            },
            LexToken::QuestionMark => {
                self.advance();
                Modifier::Conditional
            },
            _ => Modifier::None,
        }
    }

    /// 解析表达式（支持运算符优先级）
    fn parse_expression(&mut self) -> Expression {
        // 检查是否是 Excel 公式
        if matches!(self.peek(), LexToken::Identifier(_)) {
            // 检查是否是聚合函数
            if self.is_aggregate_func() {
                return self.parse_aggregate();
            }
        }

        // 使用递归下降处理优先级
        self.parse_expr()
    }

    /// 解析表达式（最低优先级：+ 和 -）
    fn parse_expr(&mut self) -> Expression {
        let mut left = self.parse_term();

        while let Some(LexToken::Operator(op)) = self.peek_op_if_add_sub() {
            self.advance();
            let right = self.parse_term();
            left = Expression::BinaryOp {
                left: Box::new(left),
                op,
                right: Box::new(right),
            };
        }

        left
    }

    /// 解析项（中等优先级：* 和 /）
    fn parse_term(&mut self) -> Expression {
        let mut left = self.parse_factor();

        while let Some(LexToken::Operator(op)) = self.peek_op_if_mul_div() {
            self.advance();
            let right = self.parse_factor();
            left = Expression::BinaryOp {
                left: Box::new(left),
                op,
                right: Box::new(right),
            };
        }

        left
    }

    /// 解析因子（最高优先级：原子表达式）
    fn parse_factor(&mut self) -> Expression {
        // 括号表达式
        if matches!(self.peek(), LexToken::ParenOpen) {
            self.advance(); // 跳过 (
            let expr = self.parse_expr();
            if matches!(self.peek(), LexToken::ParenClose) {
                self.advance(); // 跳过 )
            }
            return expr;
        }

        // 数字字面量
        if matches!(self.peek(), LexToken::Number(_)) {
            let num = self.get_number();
            self.advance();
            return Expression::Number(num);
        }

        // 标识符（变量引用或公式）
        if matches!(self.peek(), LexToken::Identifier(_)) {
            let ident = self.get_identifier();
            self.advance(); // 跳过标识符

            // 检查是否是 Excel 公式占位符 {r}
            if ident.starts_with('{') {
                // 这是公式的一部分，需要特殊处理
                // 暂时作为标识符返回
                return Expression::FieldRef(FieldRef::new(String::new(), ident));
            }

            // 检查是否是字段引用 tag.field
            if matches!(self.peek(), LexToken::Dot) {
                self.advance(); // 跳过 .
                if matches!(self.peek(), LexToken::Identifier(_)) {
                    let field = self.get_identifier();
                    self.advance();

                    let base = FieldRef::new(ident, field);

                    // 检查是否有值运算 /N 或 *N (必须是数字)
                    if matches!(self.peek(), LexToken::Operator(_)) {
                        let op = self.get_operator();
                        if matches!(op, Operator::Mul | Operator::Div) {
                            // 注意：不能提前advance，如果不是数字要留给parse_term处理
                            let next_pos = self.pos + 1;
                            if next_pos < self.tokens.len() {
                                if matches!(self.tokens[next_pos], LexToken::Number(_)) {
                                    self.advance(); // 跳过运算符
                                    let operand = self.get_number();
                                    self.advance();
                                    return Expression::ValueExpr {
                                        base,
                                        op,
                                        operand,
                                    };
                                }
                            }
                        }
                    }

                    return Expression::FieldRef(base);
                }
            }

            // 单独的标识符（可能是公式的一部分）
            return Expression::FieldRef(FieldRef::new(String::new(), ident));
        }

        // 无法解析，返回空表达式
        Expression::Number(0.0)
    }

    /// 解析聚合函数
    fn parse_aggregate(&mut self) -> Expression {
        let func_name = self.get_identifier();
        let func = AggFunc::from_str(&func_name).unwrap_or(AggFunc::Sum);
        self.advance();

        // 期望 (
        if matches!(self.peek(), LexToken::ParenOpen) {
            self.advance();
        }

        // 解析参数
        let target = if matches!(self.peek(), LexToken::Identifier(_)) {
            let tag = self.get_identifier();
            self.advance();

            if matches!(self.peek(), LexToken::Dot) {
                self.advance(); // 跳过 .
                if matches!(self.peek(), LexToken::Identifier(_)) {
                    let field = self.get_identifier();
                    self.advance();
                    AggTarget::Field(FieldRef::new(tag, field))
                } else {
                    AggTarget::Tag(tag)
                }
            } else {
                AggTarget::Tag(tag)
            }
        } else {
            AggTarget::Tag(String::new())
        };

        // 期望 )
        if matches!(self.peek(), LexToken::ParenClose) {
            self.advance();
        }

        Expression::Aggregate { func, target }
    }

    /// 解析格式码
    fn parse_format(&mut self) -> Option<FormatSpec> {
        if matches!(self.peek(), LexToken::Colon) {
            self.advance(); // 跳过 :
            if matches!(self.peek(), LexToken::FormatSpec(_)) {
                let fmt = self.get_format_spec();
                self.advance();
                return Some(FormatSpec::from_str(&fmt));
            }
        }
        None
    }

    // 辅助方法

    fn peek(&self) -> &LexToken {
        if self.pos < self.tokens.len() {
            &self.tokens[self.pos]
        } else {
            &LexToken::EOF
        }
    }

    fn advance(&mut self) {
        if !self.is_at_end() {
            self.pos += 1;
        }
    }

    fn is_at_end(&self) -> bool {
        self.pos >= self.tokens.len() || matches!(self.peek(), LexToken::EOF)
    }

    fn peek_op_if_add_sub(&self) -> Option<LexToken> {
        if self.pos < self.tokens.len() && self.tokens[self.pos].is_add_sub() {
            Some(self.tokens[self.pos].clone())
        } else {
            None
        }
    }

    fn peek_op_if_mul_div(&self) -> Option<LexToken> {
        if self.pos < self.tokens.len() && self.tokens[self.pos].is_mul_div() {
            Some(self.tokens[self.pos].clone())
        } else {
            None
        }
    }

    fn is_aggregate_func(&self) -> bool {
        if let LexToken::Identifier(name) = self.peek() {
            AggFunc::from_str(name).is_some()
        } else {
            false
        }
    }

    fn get_number(&self) -> f64 {
        if let LexToken::Number(n) = self.peek() {
            *n
        } else {
            0.0
        }
    }

    fn get_identifier(&self) -> String {
        if let LexToken::Identifier(s) = self.peek() {
            s.clone()
        } else {
            String::new()
        }
    }

    fn get_operator(&self) -> Operator {
        if let LexToken::Operator(op) = self.peek() {
            *op
        } else {
            Operator::Add
        }
    }

    fn get_format_spec(&self) -> String {
        if let LexToken::FormatSpec(s) = self.peek() {
            s.clone()
        } else {
            String::new()
        }
    }
}

/// 解析 Excel 公式中的行号占位符
pub fn replace_row_placeholders(formula: &str, row: u32) -> String {
    let mut result = formula.to_string();

    // 替换 {r}
    result = result.replace("{r}", &row.to_string());

    // 替换 {r+N}
    // 使用正则表达式更可靠
    use regex::Regex;
    let re_add = Regex::new(r"\{r\+(\d+)\}").unwrap();
    result = re_add.replace_all(&result, |caps: &regex::Captures| {
        let offset: i32 = caps.get(1).unwrap().as_str().parse().unwrap_or(0);
        (row as i32 + offset).to_string()
    }).to_string();

    // 替换 {r-N}
    let re_sub = Regex::new(r"\{r\-(\d+)\}").unwrap();
    result = re_sub.replace_all(&result, |caps: &regex::Captures| {
        let offset: i32 = caps.get(1).unwrap().as_str().parse().unwrap_or(0);
        (row as i32 - offset).to_string()
    }).to_string();

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_simple_variable() {
        let result = Parser::parse_cell("{{tag.field}}");
        assert_eq!(result.nodes.len(), 1);
        assert!(result.nodes[0].is_marker());
        assert!(matches!(result.nodes[0].get_modifier(), Some(Modifier::None)));
    }

    #[test]
    fn test_parse_row_iterate() {
        let result = Parser::parse_cell("{{@tag.field}}");
        assert!(result.is_row_iterate);
        assert!(matches!(result.nodes[0].get_modifier(), Some(Modifier::RowIterate)));
    }

    #[test]
    fn test_parse_with_format() {
        let result = Parser::parse_cell("{{tag.field:float:2}}");
        assert!(result.nodes[0].get_format().is_some());
    }

    #[test]
    fn test_parse_aggregate() {
        let result = Parser::parse_cell("{{sum(tag.field)}}");
        assert!(result.has_aggregate);
    }

    #[test]
    fn test_parse_binop() {
        let result = Parser::parse_cell("{{#(tag1.f1*tag2.f2)}}");
        assert!(matches!(result.nodes[0].get_modifier(), Some(Modifier::VariableOp)));
        let expr = result.nodes[0].get_expression().unwrap();
        // 应该是 BinaryOp
        assert!(matches!(expr, Expression::BinaryOp { .. }));
    }

    #[test]
    fn test_parse_nested_expr() {
        let result = Parser::parse_cell("{{#((tag1.f1+tag2.f2)*tag3.f3)}}");
        let expr = result.nodes[0].get_expression().unwrap();
        // 应该是嵌套的 BinaryOp
        assert!(matches!(expr, Expression::BinaryOp { .. }));
    }

    #[test]
    fn test_parse_formula() {
        let result = Parser::parse_cell("{{=A{r}*B{r}}}");
        assert!(matches!(result.nodes[0].get_modifier(), Some(Modifier::Formula)));
    }

    #[test]
    fn test_parse_row_formula() {
        let result = Parser::parse_cell("{{@=A{r}/B{r}:pct}}");
        assert!(matches!(result.nodes[0].get_modifier(), Some(Modifier::RowFormula)));
        assert!(result.is_row_iterate);
    }

    #[test]
    fn test_parse_mixed_text() {
        let result = Parser::parse_cell("Hello {{name}} World");
        assert_eq!(result.nodes.len(), 3);
        assert!(matches!(result.nodes[0], AstNode::Text(_)));
        assert!(result.nodes[1].is_marker());
        assert!(matches!(result.nodes[2], AstNode::Text(_)));
    }

    #[test]
    fn test_get_tags_from_expression() {
        let result = Parser::parse_cell("{{@limits.client}}");
        assert!(result.is_row_iterate);
        assert_eq!(result.iterate_tag, Some("limits".to_string()));
        let expr = result.nodes[0].get_expression().unwrap();
        assert_eq!(expr.get_tags(), vec!["limits".to_string()]);
    }

    #[test]
    fn test_operator_precedence() {
        // a + b * c 应该解析为 a + (b * c)
        let result = Parser::parse_cell("{{#(tag1.a+tag2.b*tag3.c)}}");
        let expr = result.nodes[0].get_expression().unwrap();

        // 顶层应该是 + 运算 (因为 * 优先级高于 +)
        if let Expression::BinaryOp { left, op, right } = expr {
            assert!(matches!(op, Operator::Add));
            // right 应该是 * 运算
            assert!(matches!(right.as_ref(), Expression::BinaryOp { .. }));
        } else {
            panic!("Expected BinaryOp");
        }
    }

    #[test]
    fn test_parse_row_formula_with_format() {
        let result = Parser::parse_cell("{{@=C{r}/B{r}:pct:3}}");
        assert!(result.is_row_iterate);
        assert!(matches!(result.nodes[0].get_modifier(), Some(Modifier::RowFormula)));

        let expr = result.nodes[0].get_expression().unwrap();
        assert!(matches!(expr, Expression::ExcelFormula { .. }));

        let format = result.nodes[0].get_format();
        assert!(format.is_some());
        assert!(matches!(format.unwrap().format_type, FormatType::Percent(3)));
    }

    #[test]
    fn test_replace_row_placeholders() {
        assert_eq!(replace_row_placeholders("A{r}", 5), "A5");
        assert_eq!(replace_row_placeholders("A{r+1}", 5), "A6");
        assert_eq!(replace_row_placeholders("A{r-1}", 5), "A4");
        assert_eq!(replace_row_placeholders("SUM(A1:A{r})", 10), "SUM(A1:A10)");
    }
}