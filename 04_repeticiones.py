# ===========================================
# Lección 4: Repeticiones (bucles / loops)
# ===========================================
#
# Un bucle repite instrucciones varias veces.
# Es como decir: "lava cada plato de la pila"

# --- Bucle 'for': repetir un número conocido de veces ---

print("Contando del 1 al 5:")
for numero in range(1, 6):
    print(numero)

# range(1, 6) genera: 1, 2, 3, 4, 5
# (el último número NO se incluye)

print()  # línea vacía para separar

# --- Recorrer una lista ---

frutas = ["manzana", "banana", "naranja"]

print("Mis frutas favoritas:")
for fruta in frutas:
    print("-", fruta)

print()

# --- Bucle 'while': repetir mientras algo sea verdadero ---

print("Cuenta regresiva:")
contador = 5
while contador > 0:
    print(contador)
    contador = contador - 1

print("¡Despegue!")

# -------------------------------------------
# Explicación:
#   for   = PARA cada elemento, haz esto
#   while = MIENTRAS esta condición sea verdadera, repite
#
# Una lista es un grupo de datos entre corchetes: [1, 2, 3]
# -------------------------------------------

# -------------------------------------------
# Ejercicio: Agrega más frutas a la lista
# y ejecuta con: python3 04_repeticiones.py
# -------------------------------------------
