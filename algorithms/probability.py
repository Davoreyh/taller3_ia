"""Punto 1. Distribuciones discretas y decisiones de respuesta."""
import math


def normalize(weights: dict) -> dict:
    """Retorna un diccionario nuevo cuyas probabilidades suman uno.

    Entrada: weights = {resultado: peso}.
    Salida: {resultado: probabilidad}, con las mismas claves.

    Tips:
    - Conserve las claves y su orden; no modifique weights.
    - Rechace pesos negativos, no finitos, diccionarios vacíos y suma cero.
    """
    # TODO: Add your code here
    if not weights:
        raise ValueError("Papi no hay datos")
    probability = {}
    total_weight = 0
    for result in weights:
        value = weights[result]
        if not math.isfinite(value) or value < 0:
            raise ValueError("Papi el valor es inválido")
        total_weight += value
    
    if total_weight == 0:
        raise ValueError("Papi no hay datos")
    
    for result in weights:
        probability[result] = weights[result]/total_weight
    
    return probability
    # FIN DEL BLOQUE DEL ESTUDIANTE


def expected_utility(distribution: dict, utilities: dict) -> float:
    """Calcula sum_s P(s) U(s) para una acción.

    Entradas: distribution = {resultado: probabilidad} y
    utilities = {resultado: utilidad} para una sola acción.
    Salida: un número con la utilidad esperada.

    Tips:
    - distribution ya debe estar normalizada; no normalice utilidades.
    - Ambas tablas deben contener los mismos resultados posibles.
    """
    # Validación de las entradas; no modificar.
    if set(distribution) != set(utilities) or not distribution:
        raise ValueError("Probabilidades y utilidades deben tener las mismas claves")
    if any(not math.isfinite(p) or p < 0 for p in distribution.values()):
        raise ValueError("Distribución inválida")
    if not math.isclose(sum(distribution.values()), 1, abs_tol=1e-9):
        raise ValueError("La distribución debe sumar uno")
    if any(not math.isfinite(u) for u in utilities.values()):
        raise ValueError("Utilidades no finitas")
    # TODO: Add your code here
    final_util = 0
    for result in distribution:
        final_util += distribution[result]*utilities[result]
    
    return final_util
    # FIN DEL BLOQUE DEL ESTUDIANTE


def best_action(distribution: dict, actions: dict) -> tuple[str, dict]:
    """Retorna (nombre de acción, utilidades esperadas por acción).

    Entradas: distribution = {resultado: probabilidad} y
    actions = {acción: {resultado: utilidad}}.
    Salida: (acción elegida, {acción: utilidad esperada}).

    Tips:
    - actions asocia cada acción con su tabla de utilidades.
    - Reutilice expected_utility y conserve la primera acción en los empates.
    """
    # Validación de las acciones; no modificar.
    if not actions:
        raise ValueError("Se requiere al menos una acción")
    # TODO: Add your code here
    final_utilities = {}
    best_action = None
    for action in actions:
        action_util = expected_utility(distribution, actions[action])
        final_utilities[action] = action_util
        if best_action is None or action_util > final_utilities[best_action]:
            best_action = action
    return best_action, final_utilities
    # FIN DEL BLOQUE DEL ESTUDIANTE
