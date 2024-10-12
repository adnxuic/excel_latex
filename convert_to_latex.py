greek_alphabet = {
    "alpha": "α",
    "beta": "β",
    "gamma": "γ",
    "delta": "δ",
    "epsilon": "ε",
    "zeta": "ζ",
    "eta": "η",
    "theta": "θ",
    "iota": "ι",
    "kappa": "κ",
    "lambda": "λ",
    "mu": "μ",
    "nu": "ν",
    "xi": "ξ",
    "omicron": "ο",
    "pi": "π",
    "rho": "ρ",
    "sigma": "σ",
    "tau": "τ",
    "upsilon": "υ",
    "phi": "φ",
    "chi": "χ",
    "psi": "ψ",
    "omega": "ω",
    "Alpha": "Α",
    "Beta": "Β",
    "Gamma": "Γ",
    "Delta": "Δ",
    "Epsilon": "Ε",
    "Zeta": "Ζ",
    "Eta": "Η",
    "Theta": "Θ",
    "Iota": "Ι",
    "Kappa": "Κ",
    "Lambda": "Λ",
    "Mu": "Μ",
    "Nu": "Ν",
    "Xi": "Ξ",
    "Omicron": "Ο",
    "Pi": "Π",
    "Rho": "Ρ",
    "Sigma": "Σ",
    "Tau": "Τ",
    "Upsilon": "Υ",
    "Phi": "Φ",
    "Chi": "Χ",
    "Psi": "Ψ",
    "Omega": "Ω"
}

class LatexConvert:
    @staticmethod
    def greek_alphabet_convert(data):
        for i, row in enumerate(data):
            for j, cell in enumerate(row):
                cell_str = str(cell)
                new_cell = []
                for char in cell_str:
                    greek_key = next((k for k, v in greek_alphabet.items() if v == char), None)
                    if greek_key:
                        new_cell.append(f"$\\{greek_key}$")
                    else:
                        new_cell.append(char)
                data[i][j] = ''.join(new_cell)
        return data

    @staticmethod
    def convert_to_latex_code(data):
        data = LatexConvert.greek_alphabet_convert(data)

        num_columns = len(data[0])
        num_rows = len(data)

        latex = "\\begin{table}[htbp]\n"
        latex += "\t\\centering\n"
        latex += "\t\t\\caption{表名}\n"
        latex += "\t\t\\vspace{1em} % 表名与表格的间距\n"
        if num_columns >= 10:
            latex += f"\t\t\\resizebox{{\\textwidth}}{{!}}\n"
        else:
            latex += f"\t\t\\resizebox{{!}}{{!}}\n"
        latex += "\t\t{\n"
        latex += f"\t\t\\begin{{tabular}}{{{'c' * num_columns}}}\n"

        latex += "\t\t\\toprule\n"

        for i, row in enumerate(data):
            latex += "\t\t" + " & ".join(str(cell) for cell in row) + " \\\\\n"
            if i == 0:
                latex += "\t\t\\midrule\n"

        latex += "\t\t\\bottomrule\n"
        latex += "\t\t\\end{tabular}\n"
        latex += "\t\t}\n"
        latex += "\t\t\\label{tab:data}\n"
        latex += "\\end{table}"
        return latex
