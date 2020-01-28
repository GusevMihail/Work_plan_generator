import pandas as pd
from typing import List, Tuple, Union, Any
from pre_processing import Job

def jobs_to_df(jobs: List[Job]) -> pd.DataFrame:
    columns = ('date', 'system', 'object', 'place', 'work_type', 'performer')
    result = pd.DataFrame(columns=columns,
                          data=((j.date, j.system, j.object, j.place, j.work_type, j.performer)
                                for j in jobs))
    return result