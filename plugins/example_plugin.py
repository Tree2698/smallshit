STAT_NAME = "离散系数(绝对中位差)"

def compute(series):
    values = series.dropna()
    if len(values) == 0:
        return None
    median = values.median()
    mad = (values - median).abs().median()
    return float(mad)
