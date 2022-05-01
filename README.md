# ComparacionTabsCeldas
Este código esta pensando para usarse como macro de Excel.

## Características:
1. Se comienza pidiendo la selección de dos archivos de excel Uno "viejo" y otro "nuevo". 
2. El programa compara el número de hojas con la que cuentan ambos archivos de excel y muestra si se tiene el mismo número de hojas o son distintas.
3. Compara el nombre de los archivos de ambos exceles y los que solo se encuentren en un solo archivo aparecerán con la pestaña de color "azul verdoso".
4. Finalmente, compara los archivos con el mismo nombre y colorea las celdas de color "magenta" aquellas celdas que sean distintas en el archivo "viejo" respecto al "nuevo". Las celdas se colorean en el archivo "nuevo".
5. Las hojas donde aparezcan los cambios mencionados en el punto 4 son coloreadas de color "negro".

### Nota:
No importa el orden de las celdas, siempre y cuando tengan el mismo nombre. Además, solo se comparan aquellas cuyo nombre de hoja comience con ZDC. 
Para cambiar este parametro, ir a la línea 60.

## A mejorar:
1. Aún falta mejorar en el caso de que no se seleccione un archivo.
2. Cuando no se selecciona un archivo sale un error y se detiene la ejecución.
