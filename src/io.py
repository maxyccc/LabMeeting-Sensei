from typing import List, Union
from glob import glob
import os


def search_pattern_file(pattern: str, search_dir: str, _only_one: bool = False) -> Union[str, List[str]]:
    recursive_path = os.path.join(search_dir, '**', pattern)
    matches = sorted(list(glob(recursive_path, recursive=True)))
    if _only_one and len(matches) != 1:
        raise ValueError(f"Expected 1 match for {pattern}, got {len(matches)}")
    return matches if not _only_one else matches[0]