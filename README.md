# Sistema MULATO EDICIONES

Sistema administrativo de la agencia: precios de los planes en dólares BCV sin perder margen, cobros del día 5, gastos de traslado, pagos al equipo en USDT y reparto del margen.

- **Página:** `index.html`, publicada con GitHub Pages.
- **Datos:** un Google Sheet privado (pestañas Config, Planes, Herramientas, Tasas, Clientes, Cobros, Gastos, Nomina).
- **Puente:** `apps-script/Code.gs`, pegado en el Apps Script del Sheet. Protegido con una clave.
- **Tasas automáticas:** el Apps Script busca cada día a las 9 am (Caracas) la tasa BCV oficial (DolarApi) y el precio de venta del USDT en Binance P2P. Si Binance no responde, usa el paralelo de DolarApi. Una tasa cargada a mano tiene prioridad.

Este repositorio es público y solo contiene código. Ningún dato de clientes, precios ni sueldos se guarda aquí.

## Cómo se calculan los precios

```
Costo del plan   = community + video maker por pauta + traslado por pauta + herramientas por cliente
Mínimo USDT      = Costo ÷ (1 − margen objetivo)
Precio $ BCV     = mayor(Mínimo, USDT que quieres recibir) × (tasa USDT ÷ tasa BCV) × (1 + colchón)
Recibes USDT     = Precio $ BCV × tasa BCV ÷ tasa USDT
```

## Instalación del Apps Script

1. Abre el Sheet > Extensiones > Apps Script. Pega `apps-script/Code.gs` y cambia `CLAVE` (y `USUARIO` si quieres otro).
2. Ejecuta la función `instalar` y acepta los permisos.
3. Implementar > Nueva implementación > Aplicación web. Ejecutar como: Yo. Acceso: Cualquier usuario.
4. Abre la página y entra con el usuario (USUARIO en el Apps Script) y tu clave. La URL ya está en `API_URL` dentro de index.html.

Si cambias el código del Apps Script: Implementar > Gestionar implementaciones > editar > Nueva versión (así la URL no cambia).

## Asesor

Pestaña **Asesor**: preguntas en lenguaje normal sobre tus números. La página envía tu pregunta y un resumen de los datos al Apps Script, y este consulta a Claude.

1. Crea una clave en console.anthropic.com (la API tiene facturación aparte de la suscripción a Claude; con 6 clientes el gasto es de centavos al mes).
2. Apps Script > Configuración del proyecto (engranaje) > Propiedades del script > Agregar: nombre `ANTHROPIC_KEY`, valor tu clave. Nunca la pegues en el código.
3. Pega el `Code.gs` nuevo, guarda y ejecuta cualquier función una vez para aceptar el permiso de conexión externa.
4. Implementar > Gestionar implementaciones > editar > Nueva versión.

El asesor ve nombres de clientes, montos y gastos de los últimos 3 meses; no ve sueldos individuales ni claves. No sustituye a un contador.
