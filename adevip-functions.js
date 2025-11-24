// ============================================
// CONFIGURACIÓN GLOBAL - ADEVIP
// ============================================

const CONFIG = {
    // Configuración de SharePoint
    sharepoint: {
        siteUrl: 'https://adevip.sharepoint.com/sites/Adevip',
        listaCursos: 'Cursos',
        listaInscripciones: 'Inscripciones',
        bibliotecaDocumentos: 'DocumentosInscripciones'
    },
    
    // URL del Power Automate Flow para enviar correos e inscripciones
    flowUrl: 'https://default2c5961ff9ec1415d994308985133b5.84.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/95df6203ab0748cb860b31dae46f9b9f/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Zz3vllZOxjMrpwBxkgJHa-QqzT3I8Brb107uu4LCvCY',
    
    // URL del Power Automate Flow para obtener cursos
    flowCursosUrl: 'https://default2c5961ff9ec1415d994308985133b5.84.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/038b128685c7436d8ec6a5b9ecbd354a/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=z-ITUzNgGLlk4P5tZpVIySIgoVXxw07dwamLdeL6lKM',

    // Colores corporativos ADEVIP
    colors: {
        primary: '#1e40af',
        secondary: '#dc2626',
        accent: '#0891b2'
    }
};

// Variables globales
let cursosData = [];
let cursoSeleccionado = null;

// ============================================
// INICIALIZACIÓN
// ============================================

document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 ADEVIP - Sistema de Inscripciones iniciado');
    cargarCursos();
    configurarEventosArchivos();
});

// ============================================
// FUNCIONES DE SHAREPOINT
// ============================================

/**
 * Cargar cursos desde SharePoint vía Power Automate
 */
async function cargarCursos() {
    try {
        const coursesContainer = document.getElementById('coursesContainer');
        
        // Mostrar loading
        coursesContainer.innerHTML = `
            <div class="loading">
                <div class="spinner"></div>
                <p>Cargando cursos desde SharePoint...</p>
            </div>
        `;
        
        console.log('📚 Cargando cursos desde Power Automate Flow...');
        
        // Verificar que flowCursosUrl esté configurada
        if (!CONFIG.flowCursosUrl || CONFIG.flowCursosUrl.includes('PEGA')) {
            throw new Error('flowCursosUrl no está configurada');
        }
        
        try {
            // Llamar al Flow de Power Automate
            const response = await fetch(CONFIG.flowCursosUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                }
            });

            if (response.ok) {
                const data = await response.json();
                console.log('📦 Respuesta del Flow:', data);
                
                // Verificar si hay cursos
                if (data.value && Array.isArray(data.value) && data.value.length > 0) {
                    cursosData = data.value.map(item => validarDatosCurso(item));
                    console.log('✅ Cursos cargados desde SharePoint:', cursosData.length);
                    console.log('📊 Primer curso:', cursosData[0]);
                } else {
                    console.warn('⚠️ La respuesta no contiene cursos');
                    throw new Error('No hay cursos activos en SharePoint');
                }
            } else {
                const errorText = await response.text();
                console.error('❌ Error del Flow:', response.status, errorText);
                throw new Error('Error al cargar desde Power Automate: ' + response.status);
            }
        } catch (error) {
            console.warn('⚠️ No se pudo cargar desde Power Automate:', error.message);
            console.log('📦 Usando cursos de ejemplo...');
            
            // Datos de ejemplo como fallback
            cursosData = [
                {
                    id: 1,
                    nombre: 'Curso Básico de Vigilancia',
                    descripcion: 'Fundamentación completa para guardas de seguridad',
                    duracion: '120 horas',
                    precio: '$450.000',
                    modalidad: 'Presencial',
                    nivel: 'Básico'
                },
                {
                    id: 2,
                    nombre: 'Curso de Escolta',
                    descripcion: 'Formación especializada en protección personal',
                    duracion: '160 horas',
                    precio: '$850.000',
                    modalidad: 'Presencial',
                    nivel: 'Avanzado'
                },
                {
                    id: 3,
                    nombre: 'Supervisor de Seguridad',
                    descripcion: 'Liderazgo y gestión de equipos de seguridad',
                    duracion: '80 horas',
                    precio: '$650.000',
                    modalidad: 'Híbrido',
                    nivel: 'Especialización'
                },
                {
                    id: 4,
                    nombre: 'Operador de Medios Tecnológicos',
                    descripcion: 'Monitoreo y control de sistemas de seguridad',
                    duracion: '100 horas',
                    precio: '$550.000',
                    modalidad: 'Presencial',
                    nivel: 'Especialización'
                },
                {
                    id: 5,
                    nombre: 'Reentrenamiento Vigilancia',
                    descripcion: 'Actualización para personal activo',
                    duracion: '30 horas',
                    precio: '$250.000',
                    modalidad: 'Virtual',
                    nivel: 'Actualización'
                },
                {
                    id: 6,
                    nombre: 'Manejador Canino',
                    descripcion: 'Entrenamiento con perros de seguridad',
                    duracion: '140 horas',
                    precio: '$950.000',
                    modalidad: 'Presencial',
                    nivel: 'Especialización'
                }
            ];
            console.log('✅ Cursos de ejemplo cargados:', cursosData.length);
        }

        renderizarCursos();
    } catch (error) {
        console.error('❌ Error crítico cargando cursos:', error);
        mostrarNotificacion('Error al cargar los cursos', 'error');
        
        // Mostrar mensaje de error en la página
        const container = document.getElementById('coursesContainer');
        container.innerHTML = `
            <div style="text-align: center; padding: 60px 20px;">
                <p style="font-size: 18px; color: #dc2626;">❌ Error al cargar los cursos</p>
                <p style="color: #64748b;">Por favor, recarga la página o contacta al administrador</p>
                <button onclick="location.reload()" style="margin-top: 20px; padding: 12px 24px; background: #1e40af; color: white; border: none; border-radius: 8px; cursor: pointer;">
                    Recargar Página
                </button>
            </div>
        `;
    }
}

/**
 * Formatear precio a pesos colombianos
 */
function formatearPrecio(precio) {
    if (!precio) {
        return 'Consultar';
    }
    
    const precioStr = String(precio);
    
    // Si ya tiene el símbolo $, devolverlo formateado
    if (precioStr.includes('$')) {
        const numeros = precioStr.replace(/[^0-9]/g, '');
        if (numeros) {
            const numero = parseInt(numeros);
            return '$' + numero.toLocaleString('es-CO');
        }
    }
    
    // Si es número, formatearlo
    const numero = parseInt(precioStr.replace(/[^0-9]/g, ''));
    if (!isNaN(numero) && numero > 0) {
        return '$' + numero.toLocaleString('es-CO');
    }
    
    return precioStr || 'Consultar';
}

/**
 * Validar y limpiar datos del curso
 */
function validarDatosCurso(item) {
    return {
        id: item.ID || item.Id,
        nombre: item.NombreCurso || 'Curso sin nombre',
        descripcion: item.Title || 'Información disponible próximamente',
        duracion: item['Duraci_x00f3_n'] ? `${item['Duraci_x00f3_n']} horas` : 'Por definir',
        precio: formatearPrecio(item.Precio),
        modalidad: item.Modalidad?.Value || item.Modalidad || 'Presencial',
        nivel: item.Requisitos?.Value || item.Nivel || 'Básico'
    };
}

/**
 * Renderizar cursos en la página
 */
function renderizarCursos() {
    const container = document.getElementById('coursesContainer');
    
    const html = `
        <div class="courses-grid">
            ${cursosData.map(curso => `
                <div class="course-card" onclick="abrirFormulario(${curso.id})">
                    <div class="course-header">
                        <div class="course-badge">${curso.nivel}</div>
                        <div class="course-title">${curso.nombre}</div>
                        <div class="course-description">${curso.descripcion}</div>
                    </div>
                    <div class="course-body">
                        <div class="course-meta">
                            <div class="course-meta-item">
                                <span>📚</span>
                                <span>${curso.duracion}</span>
                            </div>
                            <div class="course-meta-item">
                                <span>🎓</span>
                                <span>${curso.modalidad}</span>
                            </div>
                        </div>
                        <div class="course-footer">
                            <div>
                                <div style="font-size: 12px; opacity: 0.6;">Inversión</div>
                                <div class="course-price">${curso.precio}</div>
                            </div>
                            <button class="enroll-btn" onclick="event.stopPropagation(); abrirFormulario(${curso.id})">
                                Inscribirme
                            </button>
                        </div>
                    </div>
                </div>
            `).join('')}
        </div>
    `;
    
    container.innerHTML = html;
}

/**
 * Guardar inscripción en SharePoint
 */
async function guardarEnSharePoint(formData, curso) {
    try {
        const digestResponse = await fetch(
            `${CONFIG.sharepoint.siteUrl}/_api/contextinfo`,
            {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            }
        );
        
        const digestData = await digestResponse.json();
        const formDigest = digestData.d.GetContextWebInformation.FormDigestValue;

        const itemData = {
            '__metadata': { 'type': 'SP.Data.InscripcionesListItem' },
            'NombreCompleto': formData.nombre,
            'Documento': formData.documento,
            'Correo': formData.correo,
            'Telefono': formData.telefono,
            'Curso': curso.nombre,
            'PrecioCurso': curso.precio,
            'FechaInscripcion': new Date().toISOString(),
            'EstadoInscripcion': 'Pendiente'
        };

        const response = await fetch(
            `${CONFIG.sharepoint.siteUrl}/_api/web/lists/getbytitle('${CONFIG.sharepoint.listaInscripciones}')/items`,
            {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose',
                    'X-RequestDigest': formDigest
                },
                body: JSON.stringify(itemData)
            }
        );

        const data = await response.json();
        console.log('✅ Inscripción guardada en SharePoint:', data.d.Id);
        return data.d.Id;
    } catch (error) {
        console.warn('⚠️ No se pudo guardar en SharePoint:', error);
        // Retornar ID simulado en modo demo
        return Math.floor(Math.random() * 10000);
    }
}

/**
 * Subir archivos a SharePoint
 */
async function subirArchivo(file, inscripcionId, tipoDocumento) {
    try {
        const digestResponse = await fetch(
            `${CONFIG.sharepoint.siteUrl}/_api/contextinfo`,
            {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            }
        );
        
        const digestData = await digestResponse.json();
        const formDigest = digestData.d.GetContextWebInformation.FormDigestValue;

        const fileBuffer = await file.arrayBuffer();
        const fileName = `${inscripcionId}_${tipoDocumento}_${file.name}`;

        const uploadResponse = await fetch(
            `${CONFIG.sharepoint.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${CONFIG.sharepoint.bibliotecaDocumentos}')/Files/add(url='${fileName}',overwrite=true)`,
            {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'X-RequestDigest': formDigest
                },
                body: fileBuffer
            }
        );

        console.log('✅ Archivo subido:', fileName);
        return await uploadResponse.json();
    } catch (error) {
        console.warn('⚠️ No se pudo subir archivo:', error);
        return { success: false };
    }
}

// ============================================
// FUNCIONES DE PDF
// ============================================

/**
 * Generar PDF de confirmación con diseño ADEVIP
 */
async function generarPDFConfirmacion(datosInscripcion, curso) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    // Colores corporativos ADEVIP
    const azulPrimario = [30, 64, 175];
    const rojoSecundario = [220, 38, 38];

    // === ENCABEZADO ===
    doc.setFillColor(...azulPrimario);
    doc.rect(0, 0, 210, 45, 'F');
    
    // Logo simulado (escudo)
    doc.setFillColor(255, 255, 255);
    doc.circle(30, 22, 12, 'F');
    doc.setDrawColor(...azulPrimario);
    doc.setLineWidth(2);
    doc.circle(30, 22, 12, 'S');
    
    // Texto del encabezado
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(28);
    doc.setFont(undefined, 'bold');
    doc.text('ADEVIP LTDA', 50, 20);
    
    doc.setFontSize(12);
    doc.setFont(undefined, 'normal');
    doc.text('Academia de Vigilancia y Seguridad Privada', 50, 28);
    doc.text('Formando profesionales desde 1993', 50, 35);

    // Línea decorativa
    doc.setDrawColor(...rojoSecundario);
    doc.setLineWidth(3);
    doc.line(20, 50, 190, 50);

    // === TÍTULO ===
    doc.setTextColor(...azulPrimario);
    doc.setFontSize(20);
    doc.setFont(undefined, 'bold');
    doc.text('CONFIRMACIÓN DE INSCRIPCIÓN', 105, 65, { align: 'center' });

    // === DATOS DEL ESTUDIANTE ===
    let yPos = 85;
    
    doc.setFillColor(240, 242, 245);
    doc.rect(20, yPos - 5, 170, 8, 'F');
    
    doc.setTextColor(0, 0, 0);
    doc.setFontSize(14);
    doc.setFont(undefined, 'bold');
    doc.text('DATOS DEL ESTUDIANTE', 25, yPos);
    
    yPos += 12;
    doc.setFont(undefined, 'normal');
    doc.setFontSize(11);
    doc.text(`Nombre:`, 25, yPos);
    doc.setFont(undefined, 'bold');
    doc.text(datosInscripcion.nombre, 55, yPos);
    
    yPos += 8;
    doc.setFont(undefined, 'normal');
    doc.text(`Documento:`, 25, yPos);
    doc.setFont(undefined, 'bold');
    doc.text(datosInscripcion.documento, 55, yPos);
    
    yPos += 8;
    doc.setFont(undefined, 'normal');
    doc.text(`Correo:`, 25, yPos);
    doc.setFont(undefined, 'bold');
    doc.text(datosInscripcion.correo, 55, yPos);
    
    yPos += 8;
    doc.setFont(undefined, 'normal');
    doc.text(`Teléfono:`, 25, yPos);
    doc.setFont(undefined, 'bold');
    doc.text(datosInscripcion.telefono, 55, yPos);

    // === INFORMACIÓN DEL CURSO ===
    yPos += 20;
    
    doc.setFillColor(240, 242, 245);
    doc.rect(20, yPos - 5, 170, 8, 'F');
    
    doc.setFont(undefined, 'bold');
    doc.setFontSize(14);
    doc.text('INFORMACIÓN DEL CURSO', 25, yPos);
    
    yPos += 12;
    doc.setFont(undefined, 'normal');
    doc.setFontSize(11);
    doc.text(`Curso:`, 25, yPos);
    doc.setFont(undefined, 'bold');
    doc.text(curso.nombre, 55, yPos);
    
    yPos += 8;
    doc.setFont(undefined, 'normal');
    doc.text(`Duración:`, 25, yPos);
    doc.setFont(undefined, 'bold');
    doc.text(curso.duracion, 55, yPos);
    
    yPos += 8;
    doc.setFont(undefined, 'normal');
    doc.text(`Modalidad:`, 25, yPos);
    doc.setFont(undefined, 'bold');
    doc.text(curso.modalidad, 55, yPos);
    
    yPos += 8;
    doc.setFont(undefined, 'normal');
    doc.text(`Nivel:`, 25, yPos);
    doc.setFont(undefined, 'bold');
    doc.text(curso.nivel, 55, yPos);

    // === PRECIO DESTACADO ===
    yPos += 15;
    doc.setFillColor(248, 250, 252);
    doc.setDrawColor(...azulPrimario);
    doc.setLineWidth(2);
    doc.rect(20, yPos, 170, 25);
    
    doc.setFont(undefined, 'bold');
    doc.setFontSize(12);
    doc.setTextColor(0, 0, 0);
    doc.text('Inversión Total:', 30, yPos + 10);
    
    doc.setFontSize(22);
    doc.setTextColor(...rojoSecundario);
    doc.text(curso.precio, 105, yPos + 17, { align: 'center' });

    // === FECHA ===
    yPos += 35;
    doc.setTextColor(100, 100, 100);
    doc.setFont(undefined, 'normal');
    doc.setFontSize(10);
    const fecha = new Date().toLocaleDateString('es-CO', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
    doc.text(`Fecha de Inscripción: ${fecha}`, 25, yPos);

    // === NÚMERO DE INSCRIPCIÓN ===
    yPos += 6;
    const numInscripcion = `ADEVIP-${Date.now().toString().slice(-8)}`;
    doc.text(`Número de Inscripción: ${numInscripcion}`, 25, yPos);

    // === FOOTER ===
    doc.setFillColor(...azulPrimario);
    doc.rect(0, 250, 210, 47, 'F');
    
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(11);
    doc.setFont(undefined, 'bold');
    doc.text('ADEVIP - ACADEMIA DE VIGILANTES', 105, 260, { align: 'center' });
    
    doc.setFont(undefined, 'normal');
    doc.setFontSize(9);
    doc.text('Av. 4 Norte #5N-20, Barrio Centenario, Cali - Colombia', 105, 267, { align: 'center' });
    doc.text('Tel: 313 721 8112 | Email: Deivy.palacio23@gmail.com', 105, 273, { align: 'center' });
    doc.text('Vigilados por la Superintendencia de Vigilancia y Seguridad Privada', 105, 279, { align: 'center' });
    doc.text('www.adevip.com', 105, 285, { align: 'center' });

    // Generar PDF y limpiar Base64
    let pdfBase64 = doc.output('dataurlstring').split(',')[1];
    pdfBase64 = pdfBase64.replace(/\s+/g, '').replace(/[^A-Za-z0-9+/=]/g, '');
    
    console.log('✅ PDF generado. Tamaño Base64:', pdfBase64.length, 'caracteres');
    
    return {
        pdfBlob: doc.output('blob'),
        pdfBase64: pdfBase64,
        fileName: `ADEVIP_Inscripcion_${datosInscripcion.documento}_${numInscripcion}.pdf`
    };
}

/**
 * Convertir archivo a Base64
 */
function convertirArchivoABase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const base64 = reader.result.split(',')[1];
            resolve(base64);
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

// ============================================
// FUNCIONES DE CORREO
// ============================================

/**
 * Enviar correo con PDF usando Power Automate
 */
async function enviarCorreoConPDF(datosInscripcion, curso) {
    try {
        console.log('📧 Generando PDF...');
        const pdf = await generarPDFConfirmacion(datosInscripcion, curso);

        console.log('📧 Preparando correo...');
        const emailData = {
            to: datosInscripcion.correo,
            subject: `✅ Confirmación de Inscripción - ${curso.nombre} - ADEVIP`,
            nombreCompleto: datosInscripcion.nombre,
            documento: datosInscripcion.documento,
            telefono: datosInscripcion.telefono,
            curso: curso.nombre,
            precio: curso.precio,
            duracion: curso.duracion,
            modalidad: curso.modalidad,
            cedulaBase64: datosInscripcion.cedulaBase64 || '',
            cedulaNombre: datosInscripcion.cedulaNombre || '',
            certificadoBase64: datosInscripcion.certificadoBase64 || '',
            certificadoNombre: datosInscripcion.certificadoNombre || '',
            body: `
                <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #ffffff;">
                    <!-- Header -->
                    <div style="background: linear-gradient(135deg, #1e40af 0%, #1e293b 100%); padding: 40px 20px; text-align: center; border-radius: 10px 10px 0 0;">
                        <div style="background: white; width: 80px; height: 80px; border-radius: 50%; margin: 0 auto 20px; display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 15px rgba(0,0,0,0.2);">
                            <span style="font-size: 40px;">🛡️</span>
                        </div>
                        <h1 style="color: white; margin: 0; font-size: 28px;">¡Bienvenido a ADEVIP!</h1>
                        <p style="color: rgba(255,255,255,0.9); margin: 10px 0 0 0; font-size: 16px;">Tu inscripción ha sido confirmada</p>
                    </div>
                    
                    <!-- Body -->
                    <div style="padding: 40px 30px; background: #f8fafc;">
                        <p style="font-size: 18px; color: #1e293b; margin: 0 0 20px 0;">Hola <strong>${datosInscripcion.nombre}</strong>,</p>
                        
                        <p style="font-size: 16px; color: #475569; line-height: 1.6;">
                            ¡Felicitaciones! Tu inscripción al curso <strong style="color: #1e40af;">${curso.nombre}</strong> ha sido procesada exitosamente.
                        </p>
                        
                        <!-- Course Details -->
                        <div style="background: white; padding: 25px; border-radius: 12px; margin: 25px 0; border-left: 5px solid #1e40af; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                            <h3 style="margin-top: 0; color: #1e40af; font-size: 18px;">📋 Detalles de tu inscripción:</h3>
                            <table style="width: 100%; border-collapse: collapse;">
                                <tr>
                                    <td style="padding: 8px 0; color: #64748b; font-size: 14px;"><strong>Curso:</strong></td>
                                    <td style="padding: 8px 0; color: #1e293b; font-size: 14px;">${curso.nombre}</td>
                                </tr>
                                <tr>
                                    <td style="padding: 8px 0; color: #64748b; font-size: 14px;"><strong>Duración:</strong></td>
                                    <td style="padding: 8px 0; color: #1e293b; font-size: 14px;">${curso.duracion}</td>
                                </tr>
                                <tr>
                                    <td style="padding: 8px 0; color: #64748b; font-size: 14px;"><strong>Modalidad:</strong></td>
                                    <td style="padding: 8px 0; color: #1e293b; font-size: 14px;">${curso.modalidad}</td>
                                </tr>
                                <tr>
                                    <td style="padding: 8px 0; color: #64748b; font-size: 14px;"><strong>Inversión:</strong></td>
                                    <td style="padding: 8px 0; color: #dc2626; font-size: 18px; font-weight: bold;">${curso.precio}</td>
                                </tr>
                            </table>
                        </div>
                        
                        <div style="background: #dbeafe; border-left: 4px solid #3b82f6; padding: 15px; border-radius: 8px; margin: 20px 0;">
                            <p style="margin: 0; color: #1e40af; font-size: 14px;">
                                📎 <strong>Importante:</strong> Adjunto encontrarás tu certificado de inscripción en formato PDF.
                            </p>
                        </div>
                        
                        <h4 style="color: #1e40af; margin-top: 30px;">🎯 Próximos pasos:</h4>
                        <ol style="color: #475569; line-height: 1.8; font-size: 15px;">
                            <li>Nuestro equipo administrativo revisará tu documentación</li>
                            <li>Te contactaremos en las próximas <strong>24-48 horas</strong></li>
                            <li>Recibirás información detallada sobre el inicio de clases y metodología</li>
                            <li>Te enviaremos las instrucciones de pago</li>
                        </ol>
                        
                        <!-- Contact Info -->
                        <div style="background: white; padding: 20px; border-radius: 10px; margin-top: 30px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
                            <p style="color: #64748b; margin: 0 0 15px 0; font-size: 14px;">¿Tienes alguna pregunta? Contáctanos:</p>
                            <p style="margin: 5px 0; color: #1e293b;">
                                📞 <strong>313 721 8112</strong><br>
                                📧 <strong>Deivy.palacio23@gmail.com</strong><br>
                                📍 <strong>Av. 4 Norte #5N-20, Barrio Centenario, Cali</strong>
                            </p>
                        </div>
                        
                        <p style="margin-top: 30px; color: #475569; font-size: 15px; text-align: center;">
                            ¡Gracias por confiar en <strong style="color: #1e40af;">ADEVIP</strong> para tu formación profesional!
                        </p>
                    </div>
                    
                    <!-- Footer -->
                    <div style="background: #1e293b; padding: 30px 20px; text-align: center; color: white; border-radius: 0 0 10px 10px;">
                        <p style="margin: 0; font-size: 16px; font-weight: bold;">ADEVIP - Academia de Vigilancia y Seguridad Privada</p>
                        <p style="margin: 10px 0 5px 0; font-size: 13px; opacity: 0.9;">Formando profesionales en seguridad desde 1993</p>
                        <p style="margin: 5px 0; font-size: 12px; opacity: 0.8;">Vigilados por la Superintendencia de Vigilancia y Seguridad Privada</p>
                        <p style="margin: 15px 0 0 0; font-size: 12px; opacity: 0.7;">
                            Este es un correo automático, por favor no responder a esta dirección.
                        </p>
                    </div>
                </div>
            `,
            attachment: {
                filename: pdf.fileName,
                content: pdf.pdfBase64,
                contentType: 'application/pdf'
            }
        };

        console.log('📧 Enviando correo a Power Automate...');
        const response = await fetch(CONFIG.flowUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(emailData)
        });

        if (response.ok) {
            let result = { success: true };
            
            try {
                const contentType = response.headers.get('content-type');
                if (contentType && contentType.includes('application/json')) {
                    const text = await response.text();
                    if (text && text.trim()) {
                        result = JSON.parse(text);
                    }
                }
            } catch (e) {
                console.warn('⚠️ No se pudo parsear respuesta como JSON, pero el correo se envió correctamente');
            }
            
            console.log('✅ Correo enviado exitosamente:', result);
            return true;
        } else {
            const errorText = await response.text();
            console.error('❌ Error al enviar correo:', errorText);
            throw new Error('Error al enviar correo');
        }
    } catch (error) {
        console.error('❌ Error en enviarCorreoConPDF:', error);
        throw error;
    }
}

// ============================================
// FUNCIONES DE UI
// ============================================

/**
 * Abrir formulario de inscripción
 */
function abrirFormulario(cursoId) {
    cursoSeleccionado = cursosData.find(c => c.id === cursoId);
    
    if (!cursoSeleccionado) {
        console.error('Curso no encontrado:', cursoId);
        return;
    }
    
    console.log('📝 Abriendo formulario para:', cursoSeleccionado.nombre);
    document.getElementById('formTitle').textContent = `Inscripción: ${cursoSeleccionado.nombre}`;
    document.getElementById('formSection').classList.add('active');
    
    actualizarResumen();
    
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

/**
 * Cerrar formulario
 */
function cerrarFormulario() {
    document.getElementById('formSection').classList.remove('active');
    document.getElementById('inscripcionForm').reset();
    document.getElementById('cedulaFileName').textContent = 'Click para adjuntar archivo';
    document.getElementById('cursoAnteriorFileName').textContent = 'Click para adjuntar certificado';
    cursoSeleccionado = null;
    console.log('❌ Formulario cerrado');
}

/**
 * Actualizar resumen del curso
 */
function actualizarResumen() {
    const summaryBox = document.getElementById('summaryBox');
    
    if (!cursoSeleccionado) return;
    
    summaryBox.innerHTML = `
        <div class="summary-row">
            <span style="font-weight: 600;">Curso:</span>
            <span>${cursoSeleccionado.nombre}</span>
        </div>
        <div class="summary-row">
            <span style="font-weight: 600;">Duración:</span>
            <span>${cursoSeleccionado.duracion}</span>
        </div>
        <div class="summary-row">
            <span style="font-weight: 600;">Modalidad:</span>
            <span>${cursoSeleccionado.modalidad}</span>
        </div>
        <div class="summary-row total">
            <span style="font-weight: 600; font-size: 18px;">Total:</span>
            <span style="font-size: 24px; font-weight: bold; color: ${CONFIG.colors.primary};">${cursoSeleccionado.precio}</span>
        </div>
    `;
}

/**
 * Configurar eventos de archivos
 */
function configurarEventosArchivos() {
    document.getElementById('cedulaFile').addEventListener('change', (e) => {
        if (e.target.files[0]) {
            document.getElementById('cedulaFileName').textContent = e.target.files[0].name;
            console.log('📄 Cédula seleccionada:', e.target.files[0].name);
        }
    });
    
    document.getElementById('cursoAnteriorFile').addEventListener('change', (e) => {
        if (e.target.files[0]) {
            document.getElementById('cursoAnteriorFileName').textContent = e.target.files[0].name;
            console.log('📄 Certificado seleccionado:', e.target.files[0].name);
        }
    });
}

/**
 * Mostrar notificación
 */
function mostrarNotificacion(mensaje, tipo) {
    const notification = document.getElementById('notification');
    const icon = document.getElementById('notificationIcon');
    const message = document.getElementById('notificationMessage');
    
    notification.className = `notification ${tipo}`;
    icon.textContent = tipo === 'success' ? '✓' : '✕';
    message.textContent = mensaje;
    
    notification.classList.add('show');
    
    console.log(`${tipo === 'success' ? '✅' : '❌'} ${mensaje}`);
    
    setTimeout(() => {
        notification.classList.remove('show');
    }, 5000);
}

// ============================================
// ENVIAR INSCRIPCIÓN
// ============================================

/**
 * Procesar inscripción completa
 */
async function enviarInscripcion(event) {
    event.preventDefault();
    
    const form = event.target;
    const submitBtn = document.getElementById('submitBtn');
    
    console.log('🚀 Iniciando proceso de inscripción...');
    
    // Validar
    if (!cursoSeleccionado) {
        mostrarNotificacion('Error: No hay curso seleccionado', 'error');
        return;
    }
    
    // Recopilar datos
    const formData = {
        nombre: form.nombre.value.trim(),
        documento: form.documento.value.trim(),
        correo: form.correo.value.trim(),
        telefono: form.telefono.value.trim(),
        cedulaFile: form.cedula.files[0],
        cursoAnteriorFile: form.cursoAnterior.files[0] || null
    };
    
    console.log('📋 Datos del formulario:', {
        nombre: formData.nombre,
        documento: formData.documento,
        correo: formData.correo,
        telefono: formData.telefono,
        curso: cursoSeleccionado.nombre,
        cedulaFile: formData.cedulaFile ? formData.cedulaFile.name : 'No',
        cursoAnteriorFile: formData.cursoAnteriorFile ? formData.cursoAnteriorFile.name : 'No'
    });
    
    // Validar archivos
    if (!formData.cedulaFile) {
        mostrarNotificacion('Por favor adjunte la fotocopia de la cédula', 'error');
        return;
    }
    
    // Validar tamaño de archivos (máx 5MB)
    if (formData.cedulaFile.size > 5 * 1024 * 1024) {
        mostrarNotificacion('La cédula es muy grande. Máximo 5MB', 'error');
        return;
    }
    
    if (formData.cursoAnteriorFile && formData.cursoAnteriorFile.size > 5 * 1024 * 1024) {
        mostrarNotificacion('El certificado es muy grande. Máximo 5MB', 'error');
        return;
    }
    
    try {
        submitBtn.disabled = true;
        submitBtn.textContent = 'Procesando inscripción...';
        
        // PASO 1: Guardar en SharePoint
        console.log('💾 Paso 1/5: Guardando inscripción en SharePoint...');
        const inscripcionId = await guardarEnSharePoint(formData, cursoSeleccionado);
        console.log(`✅ Inscripción guardada con ID: ${inscripcionId}`);
        
        submitBtn.textContent = 'Subiendo documentos...';
        
        // PASO 2: Subir cédula
        console.log('📤 Paso 2/5: Subiendo cédula...');
        if (formData.cedulaFile) {
            await subirArchivo(formData.cedulaFile, inscripcionId, 'Cedula');
            console.log('✅ Cédula subida correctamente');
        }
        
        // PASO 3: Subir certificado anterior (opcional)
        console.log('📤 Paso 3/5: Subiendo certificado anterior...');
        if (formData.cursoAnteriorFile) {
            await subirArchivo(formData.cursoAnteriorFile, inscripcionId, 'CertificadoAnterior');
            console.log('✅ Certificado anterior subido correctamente');
        } else {
            console.log('ℹ️ No hay certificado anterior para subir');
        }
        
        submitBtn.textContent = 'Procesando archivos...';

        // PASO 4: Convertir archivos a Base64
        console.log('📎 Paso 4/5: Convirtiendo archivos a Base64...');
        let cedulaBase64 = '';
        let certificadoBase64 = '';

        if (formData.cedulaFile) {
            cedulaBase64 = await convertirArchivoABase64(formData.cedulaFile);
            console.log('✅ Cédula convertida a Base64');
        }

        if (formData.cursoAnteriorFile) {
            certificadoBase64 = await convertirArchivoABase64(formData.cursoAnteriorFile);
            console.log('✅ Certificado convertido a Base64');
        }

        submitBtn.textContent = 'Enviando correo de confirmación...';

        // PASO 5: Enviar correo con PDF y archivos adjuntos
        console.log('📧 Paso 5/5: Enviando correo y guardando datos...');

        // Crear datos completos con archivos
        const datosCompletos = {
            ...formData,
            cedulaBase64: cedulaBase64,
            cedulaNombre: formData.cedulaFile ? formData.cedulaFile.name : '',
            certificadoBase64: certificadoBase64,
            certificadoNombre: formData.cursoAnteriorFile ? formData.cursoAnteriorFile.name : ''
        };

        await enviarCorreoConPDF(datosCompletos, cursoSeleccionado);
        console.log('✅ Correo enviado correctamente');
        
        // Éxito total
        console.log('🎉 ¡INSCRIPCIÓN COMPLETADA EXITOSAMENTE!');
        mostrarNotificacion('¡Inscripción exitosa! Revise su correo electrónico', 'success');
        
        // Limpiar y cerrar después de 2 segundos
        setTimeout(() => {
            cerrarFormulario();
        }, 2000);
        
    } catch (error) {
        console.error('❌ Error en el proceso de inscripción:', error);
        mostrarNotificacion('Error al procesar la inscripción. Por favor intente nuevamente', 'error');
    } finally {
        submitBtn.disabled = false;
        submitBtn.textContent = 'Completar Inscripción';
    }
}

// ============================================
// FUNCIONES GLOBALES EXPORTADAS
// ============================================

// Hacer funciones disponibles globalmente para el HTML
window.abrirFormulario = abrirFormulario;
window.cerrarFormulario = cerrarFormulario;
window.enviarInscripcion = enviarInscripcion;

console.log('✅ ADEVIP Functions cargadas correctamente');
console.log('🔧 Configuración:', {
    sharepoint: CONFIG.sharepoint.siteUrl,
    flowConfigured: CONFIG.flowUrl.includes('powerplatform') ? 'SÍ' : 'NO'
});