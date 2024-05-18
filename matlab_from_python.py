import matlab.engine
import numpy as np

# Iniciar el motor de MATLAB
eng = matlab.engine.start_matlab()

# Crear algunos datos de ejemplo en Python (por ejemplo, un seno)
x = np.linspace(0, 2 * np.pi, 100)
y = np.sin(x)

# Convertir los datos de Python a formato MATLAB
y_matlab = matlab.double(y.tolist())

# Llamar a la función de MATLAB para crear la gráfica
eng.plot_data(y_matlab, nargout=0)

# Mantener la ventana de la gráfica abierta
input("Presiona Enter para cerrar la gráfica y terminar el programa...")

# Cerrar el motor de MATLAB
eng.quit()