"""Punto 2. Inferencia exacta por enumeración."""
from algorithms.probability import normalize


def enumeration_ask(network, query: str, evidence: dict) -> dict:
    """Calcula P(query | evidence) sobre la conjunta factorizada.

    Entradas: network es la red; query es el nombre de la
    variable consultada; evidence es {nombre de variable: 0 o 1}.
    Salida: {0: probabilidad, 1: probabilidad}.

    Tips:
    - Recorra network.variables en orden topológico.
    - Para una variable fijada multiplique su probabilidad; para una oculta
      sume las dos alternativas 0 y 1. Normalice al terminar.
    - Use copias del evento y no modifique evidence.
    - network.probability(variable, valor, evento) consulta la CPT;
      el evento debe contener los valores de todos los padres.
    - Puede usar una función auxiliar recursiva. Al terminar el recorrido,
      el producto sin factores restantes vale 1.
    - Evidencia imposible: ValueError. Consulta observada: distribución puntual
      solo si la evidencia es posible.
    """
    # Valida nombres de variables y dominio de la evidencia.
    network.validate_query(query, evidence)

    # TODO: Add your code here
    def enumerate_all(variables, event):
        if not variables:
            return 1.0
        v = variables[0]
        rest = variables[1:]
        if v in event:
            p = network.probability(v, event[v], event)
            return p * enumerate_all(rest, event)
        else:
            total = 0.0
            for val in (0, 1):
                extended = dict(event)
                extended[v] = val
                p = network.probability(v, val, extended)
                total += p * enumerate_all(rest, extended)
            return total

    if query in evidence:
        p_evidence = enumerate_all(network.variables, dict(evidence))
        if p_evidence == 0.0:
            raise ValueError("Evidencia imposible")
        val = evidence[query]
        return {val: 1.0, 1 - val: 0.0}

    distribution = {}
    for val in (0, 1):
        extended = dict(evidence)
        extended[query] = val
        distribution[val] = enumerate_all(network.variables, extended)

    if sum(distribution.values()) == 0.0:
        raise ValueError("Evidencia imposible")

    return normalize(distribution)
    # FIN DEL BLOQUE DEL ESTUDIANTE
