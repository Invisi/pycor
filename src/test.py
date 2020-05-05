import pathlib


import excel  # type: ignore


class Config:
    DELAY_SLEEP = 10


if __name__ == "__main__":
    cor = excel.Corrector.from_path(pathlib.Path(r"C:\Projects\Python\_temp\anomaly"))
    stud = excel.Student(
        pathlib.Path(
            r"C:\Projects\Python\_temp\anomaly\2020-01-02 14.53.29_xWwcwz.xlsx"
        )
    )
    real_solutions = cor.generate_solutions(stud.mat_num, stud.dummies)

    for idx, student_solution in enumerate(stud.solutions):
        corrector_solution = real_solutions[idx]
        # Make sure the student didn't somehow delete any exercise part
        for partial_idx, partial in enumerate(student_solution):
            print(
                corrector_solution[partial_idx]["name"],
                corrector_solution[partial_idx]["value"],
                partial,
            )

    print()
