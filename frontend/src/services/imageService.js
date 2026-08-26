/**
 * Aplica filtros de mejora (brillo, contraste, enfoque/sharpen) a una imagen cargada usando HTML5 Canvas.
 * @param {HTMLImageElement|Image} imgElement - Imagen original.
 * @param {Object} options - Opciones de ajuste ({ brightness, contrast, sharpen }).
 * @returns {string} - Data URL en formato base64 JPEG de la imagen procesada.
 */
export const enhanceImageCanvas = (imgElement, options = {}) => {
  const {
    brightness = 1.1,  // 1.0 es normal, >1.0 más claro
    contrast = 1.25,   // 1.0 es normal, >1.0 más contraste
    sharpen = true     // Aplica filtro de enfoque/unsharp mask
  } = options;

  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');

  canvas.width = imgElement.naturalWidth || imgElement.width;
  canvas.height = imgElement.naturalHeight || imgElement.height;

  // Dibujar imagen original
  ctx.drawImage(imgElement, 0, 0);

  // 1. Aplicar ajuste de Brillo y Contraste mediante ImageData
  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const data = imageData.data;

  // Factor de contraste
  const factor = (259 * (contrast * 255 + 255)) / (255 * (259 - contrast * 255));

  for (let i = 0; i < data.length; i += 4) {
    // Brillo
    let r = data[i] * brightness;
    let g = data[i + 1] * brightness;
    let b = data[i + 2] * brightness;

    // Contraste
    r = factor * (r - 128) + 128;
    g = factor * (g - 128) + 128;
    b = factor * (b - 128) + 128;

    // Limitar valores entre 0 y 255
    data[i] = Math.min(255, Math.max(0, r));
    data[i + 1] = Math.min(255, Math.max(0, g));
    data[i + 2] = Math.min(255, Math.max(0, b));
  }

  ctx.putImageData(imageData, 0, 0);

  // 2. Aplicar filtro de Enfoque (Kernel Convolucional / Sharpen Matrix)
  if (sharpen) {
    const sharpenedImageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const src = imageData.data;
    const dst = sharpenedImageData.data;
    const w = canvas.width;
    const h = canvas.height;

    // Matriz de Enfoque (Unsharp mask suave)
    const weights = [
       0, -1,  0,
      -1,  5, -1,
       0, -1,  0
    ];

    for (let y = 1; y < h - 1; y++) {
      for (let x = 1; x < w - 1; x++) {
        let sy = y;
        let sx = x;
        let dstOff = (y * w + x) * 4;
        let r = 0, g = 0, b = 0;

        for (let cy = 0; cy < 3; cy++) {
          for (let cx = 0; cx < 3; cx++) {
            let scy = sy + cy - 1;
            let scx = sx + cx - 1;
            let srcOff = (scy * w + scx) * 4;
            let wt = weights[cy * 3 + cx];

            r += src[srcOff] * wt;
            g += src[srcOff + 1] * wt;
            b += src[srcOff + 2] * wt;
          }
        }

        dst[dstOff] = Math.min(255, Math.max(0, r));
        dst[dstOff + 1] = Math.min(255, Math.max(0, g));
        dst[dstOff + 2] = Math.min(255, Math.max(0, b));
      }
    }

    ctx.putImageData(sharpenedImageData, 0, 0);
  }

  return canvas.toDataURL('image/jpeg', 0.92);
};

