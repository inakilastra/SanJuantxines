# Concurso de Putxeras · San Juantxines Basauri

Sitio web moderno, elegante y optimizado para móviles para seguir en directo el Concurso de Putxeras de San Juantxines (Basauri), alimentado en vivo desde Google Sheets.

---

## 🚀 Pasos para ponerlo en marcha en GitHub Pages

### 1. Actualizar tu Google Apps Script
1. Abre tu hoja de Google Sheets y ve a **Extensiones** > **Apps Script**.
2. Reemplaza el contenido de tu archivo de código por el que está en [Codigo.gs](file:///c:/Users/inaki/OneDrive/Documentos/SanJuantxines/Codigo.gs).
3. Haz clic en **Implementar** (botón azul arriba a la derecha) > **Nueva implementación**.
4. En tipo de implementación, elige **Aplicación web**.
5. Configúralo así:
   * **Descripción**: `API En Vivo Putxeras`
   * **Ejecutar como**: `Yo` (tu cuenta de Google)
   * **Quién tiene acceso**: `Cualquier usuario` (incluso anónimo)
6. Haz clic en **Implementar** y **copia la URL de la aplicación web** (tiene una forma similar a `https://script.google.com/macros/s/AKfycb.../exec`).

---

### 2. Publicar en GitHub Pages
1. Sube este repositorio a tu cuenta de GitHub (el archivo principal es [index.html](file:///c:/Users/inaki/OneDrive/Documentos/SanJuantxines/index.html)).
2. En tu repositorio de GitHub, ve a **Settings** > **Pages**.
3. En **Branch**, selecciona `main` (o `master`) y la carpeta `/ (root)`.
4. Haz clic en **Save**. En 1 o 2 minutos tendrás tu enlace público de GitHub Pages (ej. `https://tu-usuario.github.io/SanJuantxines/`).

---

### 3. Conectar la Web a tus Datos en Vivo
1. Abre tu web en GitHub Pages o ábrela en local en tu navegador haciendo doble clic en `index.html`.
2. En la esquina superior derecha, haz clic en el icono de **Ajustes** (⚙️) o en el enlace **"Configurar enlace"** del pie de página.
3. Pega la URL de tu Google Apps Script que copiaste en el paso 1 y pulsa **Guardar y Conectar**.
4. ¡Listo! La web sincronizará automáticamente cada 45 segundos o cuando pulses el botón **Actualizar**.

---

## 🎨 Características del nuevo diseño
* **Mobile-First**: Perfecto para que los participantes consulten sus notas y puestos desde el móvil en la plaza.
* **Fases automáticas sincronizadas con Google Sheet**:
  * **Inscripción**: Muestra el total de inscritos, desglose por pueblos, buscador en tiempo real y el balance económico del sorteo.
  * **Votación**: Cata a ciegas con los puestos ordenados por número de putxera para preservar el anonimato.
  * **Clasificación**: Podio interactivo (1º, 2º y 3º puesto), mención de honor y diploma para el **Premio Local Basauri**, y tabla completa con desglose de notas (*Sabor, Punto y Presencia*) con buscador instantáneo.
* **Pestañas temáticas**: Acceso rápido al *Txokito de la Transparencia* y a la *Participación por Pueblos*.
* **Modo Demostración / Simulación**: Permite previsualizar cómo luce cada fase del concurso con un solo clic.

