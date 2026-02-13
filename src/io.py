from typing import List
from glob import glob
import os


def search_pattern_file(pattern: str, search_dir: str) -> List[str]:
    recursive_path = os.path.join(search_dir, '**', pattern)
    return list(glob(recursive_path, recursive=True))