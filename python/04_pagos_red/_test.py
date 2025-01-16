from collections import Counter

# Lista base
lista_base = [1, 2, 3, 4, 5, 2, 3, 3, 4, 5, 5]

# Lista de valores que deseas validar
valores_a_validar = [2, 3, 5, 6]

# Contar la frecuencia de cada valor en la lista base
contador = Counter(lista_base)

# Verificar cuántas veces se repiten los valores de la lista de validación
resultados = {valor: contador.get(valor) for valor in valores_a_validar}
print(contador)
print("Frecuencias:", resultados)
print(contador.get(4))



# def validate_previous_counter(self, value: str, year: int, index: int) -> int:
#         current_radicado_year = value[:4]
#         if current_radicado_year != str(year):
#             # Cargamos todas las hojas relevantes una sola vez
#             df_by_year = values_validation.load_all_sheets()
#             previous_dfs = [
#                 df_by_year[str(y)] for y in range(int(current_radicado_year), year + 1)
#             ]
#             entire_df: pd.DataFrame = pd.concat(previous_dfs, ignore_index=True)
#             entire_list: list[str] = (
#                 entire_df[entire_df.columns[index]].dropna().astype(str).to_list()
#             )
#             # Number of repetitions in previous file
#             counter: int = entire_list.count(value)
#             print(f"Radicados {value} esta: ", counter)
#             return counter
#         return 0