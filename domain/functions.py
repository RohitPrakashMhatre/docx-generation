from typing import Dict, Callable

def fn_sum(values):
    return sum(v for v in values if v is not None)

def fn_avg(values):
    return sum(values) / len(values) if values else 0

FUNCTION_REGISTRY: Dict[str, Callable] = {
    "sum": fn_sum,
    "avg": fn_avg
}

def evaluate_function_placeholder(ph, data):
    """
    ph.key = 'a,b'
    ph.rule.function = 'sum'
    """
    args = [k.strip() for k in ph["key"].split(",")]
    values = [data.get(arg) for arg in args]

    func_name = ph["rules"].get("function")
    if not func_name:
        return ""

    func = FUNCTION_REGISTRY.get(func_name)
    if not func:
        raise ValueError(f"Unknown function: {func_name}")

    return str(func(values))
