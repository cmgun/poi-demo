package com.cmgun.excel.expression;

import org.apache.commons.jexl2.Expression;

/**
 * jexl表达式
 *
 * @author chenqilin
 * @Date 2019/6/17
 */
public class JexlExpression {

    /**
     * 模板中的原始表达式，包括占位符，即 ${....}
     */
    private String originExpression;

    private Expression expression;

    public JexlExpression(String originExpression, Expression expression) {
        this.originExpression = originExpression;
        this.expression = expression;
    }

    public String getOriginExpression() {
        return originExpression;
    }

    public void setOriginExpression(String originExpression) {
        this.originExpression = originExpression;
    }

    public Expression getExpression() {
        return expression;
    }

    public void setExpression(Expression expression) {
        this.expression = expression;
    }
}
