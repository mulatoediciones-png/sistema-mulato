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

1. Abre el Sheet > Extensiones > Apps Script. Pega `apps-script/Code.gs` y cambia `CLAVE`.
2. Ejecuta la función `instalar` y acepta los permisos.
3. Implementar > Nueva implementación > Aplicación web. Ejecutar como: Yo. Acceso: Cualquier usuario.
4. Abre la página, pega la URL de la aplicación web y la clave.

Si cambias el código del Apps Script: Implementar > Gestionar implementaciones > editar > Nueva versión (así la URL no cambia).
