// @author Chris Demmings - https://demmings.github.io/
/*  *** DEBUG START ***
//  Remove comments for testing in NODE
import { SqlParse } from './SimpleParser.js';
export { Select2Query };

class Logger {
    static log(msg) {
        console.log(msg);
    }
}
//  *** DEBUG END ***/


class Select2Query {
    setTables(...parms) {
        this.tables = this.addTableData(parms);
        return this;
    }

    setTableMap(tables) {
        this.tables = tables;
        return this;
    }

    convert(selectStatement) {
        let queryStatement = "";
        let ast = null;
        let query = "";

        if (typeof selectStatement === 'string') {
            try {
                ast = SqlParse.sql2ast(selectStatement);
            }
            catch (ex) {
                return ex.toString();
            }
        }
        else {
            ast = selectStatement;
        }

        if (typeof ast.JOIN !== 'undefined') {
            let joinQuery = new QueryJoin(this.tables);
            query = joinQuery.join(ast);
        }
        else {
            queryStatement = this.selectFields(ast);
            queryStatement += this.whereCondition(ast);
            queryStatement += this.groupBy(ast);
            queryStatement += this.orderBy(ast);

            query = this.formatAsQuery(queryStatement, ast.FROM.table);
        }

        console.log(query);

        return query;
    }

    formatAsQuery(statement, tableName) {
        let range = this.tables.has(tableName.toUpperCase()) ? this.tables.get(tableName.toUpperCase()) : "";

        return "=QUERY(" + range + ", \"" + statement + "\")";
    }

    addTableData(parms) {
        const tables = new Map();

        //  Should be:  TABLE NAME, TABLE RANGE, name, range, name, range,...
        let i = 0;
        while (i + 1 < parms.length) {
            console.log(`Add Table: ${parms[i]}. Items=${parms[i + 1].length}`);
            tables.set(parms[i].trim().toUpperCase(), parms[i + 1]);
            i += 2;
        }

        return tables;
    }

    selectFields(ast) {
        let selectStr = "SELECT ";
        if (typeof ast.SELECT !== 'undefined') {
            for (let i = 0; i < ast.SELECT.length; i++) {
                let fld = ast.SELECT[i];

                let fieldName = fld.name;
                if (fld.name.indexOf(".") !== -1) {
                    let parts = fld.name.split(".");
                    fieldName = parts[1];
                }
                selectStr += fieldName.toUpperCase();

                if (i + 1 < ast.SELECT.length) {
                    selectStr += ", ";
                }
            }
        }

        return selectStr;
    }

    /**
  * Retrieve filtered record ID's.
  * @param {Object} ast - Abstract Syntax Tree
  * @returns {String} - Query WHERE condition.
  */
    whereCondition(ast) {
        let queryWhere = "";

        let conditions = {};
        if (typeof ast.WHERE !== 'undefined') {
            conditions = ast.WHERE;
        }
        else {
            //  Entire table is selected.  
            return queryWhere;
        }

        if (typeof conditions.logic === 'undefined')
            queryWhere = this.resolveCondition("OR", [conditions], "");
        else
            queryWhere = this.resolveCondition(conditions.logic, conditions.terms, "");

        return " WHERE " + queryWhere.trim();
    }

    /**
    * Recursively resolve WHERE condition and then apply AND/OR logic to results.
    * @param {String} logic - logic condition (AND/OR) between terms
    * @param {Object} terms - terms of WHERE condition (value compared to value)
    * @returns {String} - condition of where 
    */
    resolveCondition(logic, terms, queryWhere) {

        for (let i = 0; i < terms.length; i++) {
            let cond = terms[i];

            if (typeof cond.logic === 'undefined') {
                queryWhere += this.whereString(cond);
            }
            else {
                queryWhere = this.resolveCondition(cond.logic, cond.terms, queryWhere);
            }

            if (i + 1 < terms.length) {
                queryWhere += " " + logic;
            }
        }

        return queryWhere;
    }

    whereString(cond) {
        let whereStr = "";

        switch (cond.operator) {
            case "NOT IN":
                whereStr = " NOT " + this.whereValue(cond.left) + " " + this.whereOp("IN") + " " + this.whereValue(cond.right);
                break;
            case "IN":
            default:
                whereStr = " " + this.whereValue(cond.left) + " " + this.whereOp(cond.operator) + " " + this.whereValue(cond.right);
                break;
        }

        return whereStr;
    }

    whereOp(operator) {
        if (operator.trim().toUpperCase() === 'IN')
            return "MATCHES";

        return operator;
    }

    whereValue(cond) {
        if (typeof cond.SELECT === 'undefined') {
            return cond;
        }
        else {
            const sql = new Select2Query().setTableMap(this.tables)

            const preTextJoin = "'\"&TEXTJOIN(\"|\", true, ";
            const postTextJoin = ")&\"'";
            const query = sql.convert(cond).slice(1);       //  Get rid of starting '='.
            return preTextJoin + query + postTextJoin;
        }
    }

    orderBy(ast) {
        let orderBy = "";

        if (typeof ast['ORDER BY'] === 'undefined') {
            return orderBy;
        }

        orderBy = " ORDER BY "
        for (let i = 0; i < ast['ORDER BY'].length; i++) {
            let order = ast['ORDER BY'][i];
            orderBy += order.name + " " + order.order.toUpperCase();

            if (i + 1 < ast['ORDER BY'].length) {
                orderBy += ", ";
            }
        }

        return orderBy;
    }

    groupBy(ast) {
        let groupBy = "";

        if (typeof ast['GROUP BY'] === 'undefined') {
            return groupBy;
        }

        groupBy += " GROUP BY ";

        for (let i = 0; i < ast['GROUP BY'].length; i++) {
            let order = ast['GROUP BY'][i];
            groupBy += order.name;

            if (i + 1 < ast['GROUP BY'].length) {
                groupBy += ", ";
            }
        }

        return groupBy;
    }


}

class QueryJoin {
    constructor(tables) {
        this.tables = tables;
    }

    join(ast) {
        let query = "";

        for (let joinAst of ast.JOIN) {
            query += this.joinCondition(ast.FROM.table, ast.SELECT, joinAst, ast);
        }

        return query;
    }

    /**
     * @param {String} leftTable
     * @param {Object} selectFields
     * @param {Object} joinAst 
     * @param {Object} ast
     * @returns {String}
     */
    joinCondition(leftTable, selectFields, joinAst, ast) {
        const LEFT_KEY_RANGE = "$$LEFT_KEY$$";
        const LEFT_SELECT_FIELDS = "$$LEFT_SELECT$$";
        const RIGHT_KEY_RANGE = "$$RIGHT_KEY$$";
        const RIGHT_SELECT_FIELDS = "$$RIGHT_SELECT$$";
        const NO_MATCH_QUERY = "$$NO_MATCH$$";

        let query = "";
        let rightTable = joinAst.table;
        let conditionLeft = joinAst.cond.left;
        let conditionRight = joinAst.cond.right

        if (joinAst.type === 'right') {
            let temp = leftTable;
            leftTable = rightTable;
            rightTable = temp;

            temp = conditionLeft;
            conditionLeft = conditionRight;
            conditionRight = temp;
        }

        const leftKeyRangeValue = this.getKeyRangeString(leftTable, conditionLeft);
        const rightKeyRangeValue = this.getKeyRangeString(rightTable, conditionRight);
        const leftSelectFieldValue = this.leftSelectFields(selectFields, leftTable);
        let rightSelectFieldValue = this.rightSelectFields(selectFields, rightTable);
        const notFoundQuery = this.selectNotInJoin(ast, joinAst, leftTable, rightTable);

        if (leftSelectFieldValue !== '' && rightSelectFieldValue !== '') {
            rightSelectFieldValue = '&"!"&' + rightSelectFieldValue;
        }

        const joinFormatString = '={ArrayFormula(Split(Query(Flatten(IF($$LEFT_KEY$$=Split(Textjoin("!",1,$$RIGHT_KEY$$),"!"),$$LEFT_SELECT$$ $$RIGHT_SELECT$$,)),"Where Col1!=\'\'"),"!"))$$NO_MATCH$$}';

        query = joinFormatString.replace(LEFT_KEY_RANGE, leftKeyRangeValue);
        query = query.replace(RIGHT_KEY_RANGE, rightKeyRangeValue);
        query = query.replace(LEFT_SELECT_FIELDS, leftSelectFieldValue);
        query = query.replace(RIGHT_SELECT_FIELDS, rightSelectFieldValue);
        query = query.replace(NO_MATCH_QUERY, notFoundQuery);

        const sql = new Select2Query();
        sql.setTableMap(this.tables);

        return query;
    }

    /**
     * 
     * @param {String} tableName 
     * @param {String} condition 
     * @returns {String}
     */
    getKeyRangeString(tableName, condition) {
        let tableInfo = this.tables.get(tableName.toUpperCase());
        let field = condition;
        if (condition.indexOf(".") !== -1) {
            let parts = condition.split(".");
            field = parts[1];
        }
        let range = tableInfo;
        let rangeTable = "";
        if (range.indexOf("!") !== -1) {
            let parts = range.split("!");
            rangeTable = parts[0] + "!";
            range = parts[1];
        }

        let rangeComponents = range.split(":");
        let startRange = this.replaceColumn(rangeComponents[0], field);
        let endRange = this.replaceColumn(rangeComponents[1], field);

        return rangeTable + startRange + ":" + endRange;
    }

    replaceColumn(rowColumn, newColumn) {
        let num = rowColumn.replace(/\D/g, '');
        return newColumn + num;
    }

    selectNotInJoin(ast, joinAst, leftTable, rightTable) {
        if (joinAst.type === 'inner') {
            return "";
        }

        leftTable = leftTable.toUpperCase();

        //  Sort fields so that LEFT table are first and RIGHT table are after.
        const sortedFields = this.sortSelectJoinFields(ast, leftTable);

        const selectFlds = this.createSelectFieldsString(sortedFields);
        const label = this.createSelectLabelString(sortedFields);
        const rightFieldName = this.getJoinField(joinAst, joinAst.cond.right, joinAst.cond.left);
        const leftFieldName = this.getJoinField(joinAst, joinAst.cond.left, joinAst.cond.right);
        const leftRange = this.tables.get(leftTable.toUpperCase());
        const rightRange = this.tables.get(rightTable.toUpperCase());

        const queryStart = "QUERY(" + leftRange + ",";
        let selectStr = queryStart + "\"select " + selectFlds + " where " + rightFieldName + " is not null and NOT " + rightFieldName + " MATCHES ";

        const matchesQuery = '\'"&TEXTJOIN("|", true, QUERY(' + rightRange + ', "SELECT ' + leftFieldName + ' where ' + leftFieldName + ' is not null"))&';
        selectStr += matchesQuery;
        selectStr = selectStr + label + '", 0)';

        //  If no records are found, we need to insert an empty record - otherwise we get an array error.
        selectStr = ";IFNA(" + selectStr + "," + this.ifNaResult(ast) + ")";


        return selectStr;
    }

    /**
     * 
     * @param {Object} joinAst 
     * @param {String} rightSide 
     * @param {String} leftSide 
     * @returns {String}
     */
    getJoinField(joinAst, rightSide, leftSide) {
        let fieldName = "";
        if (joinAst.type === 'right') {
            fieldName = rightSide;
        }
        else {
            fieldName = leftSide;
        }
        if (fieldName.indexOf(".") !== -1) {
            let parts = fieldName.split(".");
            fieldName = parts[1];
        }

        return fieldName;
    }

    /**
     * @typedef {Object} JoinSelectField
     * @property {String} fieldTable
     * @property {String} fieldName
     * @property {Boolean} isNull
     */

    /**
     * @param {Object} ast 
     * @param {String} leftTable
     * @returns {JoinSelectField[]}
     */
    sortSelectJoinFields(ast, leftTable) {
        //  Sort fields so that LEFT table are first and RIGHT table are after.
        let leftFields = [];
        let rightFields = [];
        let nullCnt = 1;

        for (let fld of ast.SELECT) {
            let fieldTable = leftTable.toUpperCase();
            let fieldName = fld.name.toUpperCase();
            if (fld.name.indexOf(".") !== -1) {
                let parts = fld.name.split(".");
                fieldTable = parts[0].toUpperCase();
                fieldName = parts[1].toUpperCase();
            }

            let isNull = false;
            let fieldInfo = {
                fieldTable,
                fieldName,
                isNull
            };

            if (fieldTable === leftTable) {
                leftFields.push(fieldInfo);
            }
            else {
                fieldInfo.fieldName = "'null" + " ".repeat(nullCnt) + "'";
                fieldInfo.isNull = true;
                nullCnt++;
                rightFields.push(fieldInfo);
            }
        }

        return leftFields.concat(rightFields);
    }

    /**
     * Build a string that will be used in IFNA() when select does not find any records.
     * @param {Object} ast 
     * @returns {String}
     */
    ifNaResult(ast) {
        let naResult = "";
        for (const element of ast.SELECT) {
            naResult = (naResult === '') ? "" : naResult + ",";
            naResult += '""';
        }

        if (naResult !== "") {
            naResult = "{" + naResult + "}";
        }

        return naResult;
    }

    /**
     * 
     * @param {JoinSelectField[]} sortedFields 
     * @returns {String}
     */
    createSelectFieldsString(sortedFields) {
        let selectFlds = "";

        for (let fld of sortedFields) {
            selectFlds = (selectFlds === '' ? '' : selectFlds + ",") + fld.fieldName;
        }

        return selectFlds;
    }

    /**
    * 
    * @param {JoinSelectField[]} sortedFields 
    * @returns {String}
    */
    createSelectLabelString(sortedFields) {
        let label = "";

        for (let fld of sortedFields) {
            if (fld.isNull) {
                label = label !== "" ? label += ", " : "";
                label += fld.fieldName + " ''";
            }
        }

        if (label !== "") {
            label = "\"' label " + label;
        }

        return label;
    }

    /**
     * Assembled selected fields string for LEFT.
     * @param {Object} ast 
     * @param {String} leftTable 
     * @returns {String}
     */
    leftSelectFields(ast, leftTable) {
        let leftSelect = "";

        for (let fld of ast) {
            let selectField = "";

            if (fld.name.indexOf(".") === -1) {
                selectField = fld.name;
            }
            else {
                let parts = fld.name.split(".");
                if (parts[0].toUpperCase() === leftTable.toUpperCase()) {
                    selectField = parts[1];
                }
            }

            let rangeTable = "";
            let range = "";
            if (selectField !== "") {
                const tableInfo = this.tables.get(leftTable.toUpperCase());

                if (tableInfo.indexOf("!") !== -1) {
                    const parts = tableInfo.split("!");
                    rangeTable = parts[0] + "!";
                    range = parts[1];
                }

                const rangeComponents = range.split(":");
                const startRange = this.replaceColumn(rangeComponents[0], selectField);
                const endRange = this.replaceColumn(rangeComponents[1], selectField);

                selectField = rangeTable + startRange + ":" + endRange;

                leftSelect = leftSelect === '' ? '' : leftSelect + '&"!"& ';

                leftSelect += 'IF(' + selectField + ' <> "",' + selectField + ', " ")';
            }
        }

        return leftSelect;
    }

    /**
     * assemble SELECTED FIELDS string for RIGHT.
     * @param {Object} ast 
     * @param {String} rightTable 
     * @returns {String}
     */
    rightSelectFields(ast, rightTable) {
        let rightSelect = "";

        for (let fld of ast) {
            let selectField = "";

            if (fld.name.indexOf(".") === -1) {
                selectField = fld.name;
            }
            else {
                const parts = fld.name.split(".");
                if (parts[0].toUpperCase() === rightTable.toUpperCase()) {
                    selectField = parts[1];
                }
            }

            let rangeTable = "";
            let range = "";
            if (selectField !== "") {
                const tableInfo = this.tables.get(rightTable.toUpperCase());

                if (tableInfo.indexOf("!") !== -1) {
                    const parts = tableInfo.split("!");
                    rangeTable = parts[0] + "!";
                    range = parts[1];
                }

                const rangeComponents = range.split(":");
                const startRange = this.replaceColumn(rangeComponents[0], selectField);
                const endRange = this.replaceColumn(rangeComponents[1], selectField);

                selectField = rangeTable + startRange + ":" + endRange;

                rightSelect = rightSelect === '' ? '' : rightSelect + '&"!"& ';

                rightSelect += 'Split(Textjoin("!",1,IF(' + selectField + '<>"",' + selectField + '," ")),"!")';
            }
        }

        return rightSelect;
    }
}