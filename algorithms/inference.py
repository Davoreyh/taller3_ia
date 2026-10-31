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
    def recursive_enumerate(network, i: int, event: dict)-> float:
      if i == len(network.variables):
        return 1.0
      var = network.variables[i]
      if var in event:
        return network.probability(var, event[var], event)* recursive_enumerate(network, i+1, event)
      else:
        result = 0
        for j in range(2):
          new_event = event.copy()
          new_event[var] = j
          result += network.probability(var, j, event)*recursive_enumerate(network, i+1, new_event)
        return result
      
    possible_evidence = recursive_enumerate(network, 0, evidence)
    if not possible_evidence:
      raise ValueError("La evidencia es imposible")
    results = {}
    if query in evidence:
      if evidence[query] == 0:
        results[0] = 1.0
        results[1] = 0.0
      else:
        results[0] = 0.0
        results[1] = 1.0
    else:
      for k in range(2):
        new_evidence = evidence.copy()
        new_evidence[query] = k
        results[k] = recursive_enumerate(network, 0, new_evidence)
    return normalize(results)
    # FIN DEL BLOQUE DEL ESTUDIANTE
