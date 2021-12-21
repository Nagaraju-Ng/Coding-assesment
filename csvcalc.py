import operator

import pandas as pd
import re


class EvaluateExcel:
    def __init__(self, path):
        self.operators = {"+": operator.add,
                          "-": operator.sub,
                          "*": operator.mul,
                          "/": operator.truediv,
                          "%": operator.mod}
        self.path = path
        self.df = self.read_csv()

    def read_csv(self):
        df = pd.read_csv(self.path, index_col="i")
        return df

    def get_value(self, string):
        column = string[0]
        row = string[1:]
        value = self.df.get(column)[int(row)]
        if type(value) == str and value.startswith("="):
            value = self.evaluate(value)
        return int(value)

    def evaluate_off(self, string):
        match = re.match(r"\((.*)([-+*/])(.*)\)", string)
        lhs = match.group(1)
        lhs = int(lhs) if lhs.isdigit() else self.get_value(lhs)
        opr = self.operators[match.group(2)]
        rhs = match.group(3)
        rhs = int(rhs) if rhs.isdigit() else self.get_value(rhs)

        return opr(lhs, rhs)

    def evaluate_sum(self, string):
        match = re.match(r"sum\((.*):(.*)\)", string)
        column_lhs = match.group(1)[0]
        column_rhs = match.group(2)[0]
        lower_limit = int(match.group(1)[1:])
        upper_limit = int(match.group(2)[1:])
        result_list = self.df[column_lhs].tolist()[lower_limit - 1: upper_limit]
        if column_lhs != column_rhs:
            result_list += self.df[column_rhs].tolist()[lower_limit - 1: upper_limit]
        return sum(result_list)

    def evaluate_min(self, string):
        match = re.match(r"min\((.*):(.*)\)", string)
        column_lhs = match.group(1)[0]
        column_rhs = match.group(2)[0]
        lower_limit = int(match.group(1)[1:])
        upper_limit = int(match.group(2)[1:])
        result_list = self.df[column_lhs].tolist()[lower_limit - 1: upper_limit]
        if column_lhs != column_rhs:
            result_list += self.df[column_rhs].tolist()[lower_limit - 1: upper_limit]
        return int(min(result_list))

    def evaluate_max(self, string):
        match = re.match(r"max\((.*):(.*)\)", string)
        column_lhs = match.group(1)[0]
        column_rhs = match.group(2)[0]
        lower_limit = int(match.group(1)[1:])
        upper_limit = int(match.group(2)[1:])
        result_list = self.df[column_lhs].tolist()[lower_limit - 1: upper_limit]
        if column_lhs != column_rhs:
            result_list += self.df[column_rhs].tolist()[lower_limit - 1: upper_limit]
        return int(max(result_list))

    def evaluate_exp(self, expression):
        if expression.isdigit():
            return int(expression)
        elif expression.isalnum():
            return int(self.get_value(expression))
        else:
            i = 0
            j = 0
            new_string = ""
            while i < len(expression):
                if expression[i] in self.operators.keys():
                    new_string += str(self.evaluate_exp(expression[j:i])) + str(expression[i])
                    j = i + 1
                i += 1
            new_string += str(self.evaluate_exp(expression[j:i]))
            result = eval(new_string)
            return result

    def evaluate(self, expression):

        if expression.startswith("="):
            expression = expression[1:]
        match_off = re.match(r"(.*)(\(.*[-+*/].*\))(.*)", expression)
        match_sum = re.match(r"(.*)(sum\(.*:.*\))(.*)", expression)
        match_min = re.match(r"(.*)(min\(.*:.*\))(.*)", expression)
        match_max = re.match(r"(.*)(max\(.*:.*\))(.*)", expression)
        match_opr = re.match(r"(.*[-+*/].*)", expression)

        result = None
        if match_opr and not match_max and not match_off and not match_min and not match_sum:
            string = match_opr.group(1)
            start = re.match(r"^([-+*/])([A-Z0-9]+)", string)
            end = re.match(r"([A-Z0-9]+)([-+*/])$", string)

            if start:
                return start.group(1), self.evaluate_exp(start.group(2))
            if end:
                return end.group(2), self.evaluate_exp(end.group(1))
            return self.evaluate_exp(string)
        if match_off:

            result = self.evaluate_off(match_off.group(2))
            if match_off.group(1):
                opr, value = self.evaluate(match_off.group(1))
                if opr:
                    result = self.operators[opr](result, value)
            if match_off.group(3):
                opr, value = self.evaluate(match_off.group(3))
                if opr:
                    result = self.operators[opr](result, value)
        if match_sum:
            result = self.evaluate_sum(match_sum.group(2))
            if match_sum.group(1):
                opr, value = self.evaluate(match_sum.group(1))
                if opr:
                    result = self.operators[opr](result, value)
            if match_sum.group(3):
                opr, value = self.evaluate(match_sum.group(3))
                if opr:
                    result = self.operators[opr](result, value)
        if match_min:
            result = self.evaluate_min(match_min.group(2))
            if match_min.group(1):
                opr, value = self.evaluate(match_min.group(1))
                if opr:
                    result = self.operators[opr](result, value)
            if match_min.group(3):
                opr, value = self.evaluate(match_min.group(1))
                if opr:
                    result = self.operators[opr](result, value)
        if match_max:
            result = self.evaluate_max(match_max.group(2))

            if match_max.group(1):
                opr, value = self.evaluate(match_max.group(1))
                if opr:
                    result = self.operators[opr](result, value)
            if match_max.group(3):
                opr, value = self.evaluate(match_max.group(1))
                if opr:
                    result = self.operators[opr](result, value)
        return result


if __name__ == '__main__':
    i = 0
    ee = EvaluateExcel(r"input.csv")
    df = ee.df
    for column in ee.df.columns:
        for i in range(1, len(df[column]) + 1):
            row = df[column][i]
            if type(row) == str and row.startswith("="):
                df[column].replace(row, value=ee.evaluate(row), inplace=True)
    df.to_csv(r"output.csv", index=True)
