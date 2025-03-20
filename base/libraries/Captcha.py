import easyocr
import cv2
import numpy as np
import re

def solve_captcha(captcha_image_path):
    # Intentar cargar la imagen
    img = cv2.imread(captcha_image_path)
    if img is None:
        raise ValueError(f"No se pudo cargar la imagen en la ruta: {captcha_image_path}")

    # Convertir la imagen a escala de grises
    gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Ajuste de umbral
    _, threshold_img = cv2.threshold(gray_img, 120, 255, cv2.THRESH_BINARY_INV)
    
    # Función de OCR
    def try_ocr(image):
        reader = easyocr.Reader(['en'])
        result = reader.readtext(image)
        if result:
            # Concatenar los resultados y filtrar solo los caracteres alfanuméricos
            detected_text = "".join([x[1] for x in result]).replace(" ", "")
            # Filtrar solo letras y números
            filtered_text = re.sub(r'[^A-Za-z0-9]', '', detected_text)
            return filtered_text
        return ""
    
    # Intentar OCR con la primera configuración
    captcha_result = try_ocr(threshold_img)
    
    if captcha_result:
        return captcha_result
    
    # Si el primer intento no dio resultados, se intentan más configuraciones.
    # Intentar con distintas variaciones de contraste y brillo
    for i in range(3):  # Limitar el número de intentos (puedes ajustar este número)
        # Aumentar el contraste
        alpha = 2.0 + i * 0.5  # Modificar el contraste
        img_contrast = cv2.convertScaleAbs(threshold_img, alpha=alpha, beta=0)
        
        # Aumentar el brillo
        beta = 50 + i * 10  # Modificar el brillo
        img_brightness = cv2.convertScaleAbs(img_contrast, alpha=1.0, beta=beta)
        
        # Usar el OCR en la nueva imagen procesada
        captcha_result = try_ocr(img_brightness)
        
        if captcha_result:
            return captcha_result
        
    return "Captcha no detectado"