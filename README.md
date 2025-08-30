Para que tu programa funcione correctamente en una nueva computadora usando el archivo .exe generado con PyInstaller, necesitas instalar lo siguiente antes de compilar:

Python 3.9 o superior (solo para compilar, no para ejecutar el .exe).
Las siguientes librerías de Python:
openpyxl (para exportar a Excel)
pyinstaller (para crear el .exe)
Comandos recomendados (en PowerShell):

pip install openpyxl
pip install pyinstaller

Luego, para generar el .exe:

pyinstaller --noconfirm --onefile --name ControladorStock stock_app.py

En la nueva PC, solo necesitas copiar el .exe y el archivo de base de datos (inventario.db). No hace falta instalar Python ni las librerías si solo vas a ejecutar el .exe.

¿Te gustaría un instructivo paso a paso para usuarios finales?
