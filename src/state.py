import datetime
import typing
from pathlib import Path

from pydantic import BaseModel  # type: ignore


class CorrectorDict(BaseModel):
    codename: str
    deadline: datetime.datetime
    exercise_ranges: typing.List[typing.List[int]]
    max_attempts: int
    password: str
    title: str
    change_date: datetime.datetime
    dummy_count: int = 8


class State(BaseModel):
    correctors: typing.Dict[str, CorrectorDict] = {}  # relevant path, dict

    def save(self):
        cf = Path("state.json")
        try:
            cf.write_text(self.json(sort_keys=True, indent=4))
        except OSError:
            print("Failed to write state file")
            raise

    @staticmethod
    def load():
        cf = Path("state.json")
        if cf.exists():
            return State.parse_file(cf)
        else:
            return State()
