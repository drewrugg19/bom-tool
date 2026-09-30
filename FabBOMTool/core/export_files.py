"""Reserve export names atomically so concurrent runs cannot overwrite files."""
from contextlib import contextmanager
from pathlib import Path


@contextmanager
def reserve_export(directory, filename):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    name = Path(str(filename).replace("\\", "/")).name
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    stem = name[:-5]
    number = 1
    while True:
        path = directory / (name if number == 1 else f"{stem}_{number}.xlsx")
        try:
            # Exclusive creation is the reservation, not a check-then-write.
            with path.open("xb"):
                pass
            break
        except FileExistsError:
            number += 1
    try:
        yield path
    except BaseException:
        # Only this run's reserved file is removed after a failed export.
        path.unlink(missing_ok=True)
        raise
