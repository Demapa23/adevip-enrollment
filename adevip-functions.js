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
let categoriaSeleccionada = null;

// ============================================
// INICIALIZACIÓN
// ============================================

document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 ADEVIP - Sistema de Inscripciones iniciado');
    // Mostramos las categorías al inicio
    configurarEventosArchivos();
});

// ============================================
// FUNCIONES DE NAVEGACIÓN
// ============================================

/**
 * Seleccionar categoría y cargar cursos
 */
function seleccionarCategoria(categoria) {
    categoriaSeleccionada = categoria;
    console.log('📁 Categoría seleccionada:', categoria);
    
    // Ocultar categorías
    document.getElementById('categoriesSection').style.display = 'none';
    
    // Mostrar sección de cursos
    document.getElementById('coursesSection').classList.add('active');
    
    // Actualizar títulos según categoría
    const titulos = {
        fundamentacion: 'Cursos de Fundamentación',
        reentrenamiento: 'Cursos de Reentrenamiento',
        especializacion: 'Cursos de Especialización'
    };
    
    const subtitulos = {
        fundamentacion: 'Inicia tu formación profesional en seguridad',
        reentrenamiento: 'Actualiza y renueva tus conocimientos',
        especializacion: 'Destaca con formación avanzada'
    };
    
    document.getElementById('coursesTitle').textContent = titulos[categoria];
    document.getElementById('coursesSubtitle').textContent = subtitulos[categoria];
    
    // Cargar cursos desde SharePoint
    cargarCursos(categoria);
    
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

/**
 * Volver a la selección de categorías
 */
function volverACategorias() {
    document.getElementById('coursesSection').classList.remove('active');
    document.getElementById('categoriesSection').style.display = 'block';
    categoriaSeleccionada = null;
    cursosData = [];
    console.log('← Volviendo a categorías');
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ============================================
// FUNCIONES DE SHAREPOINT
// ============================================

/**
 * Cargar cursos desde SharePoint vía Power Automate
 */
async function cargarCursos(categoria = null) {
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
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json'
                }
            });

            if (response.ok) {
                const data = await response.json();
                console.log('📦 Respuesta del Flow:', data);
                
                // Verificar si hay cursos
                if (data.value && Array.isArray(data.value) && data.value.length > 0) {
                    // Procesar todos los cursos
                    let todosCursos = data.value.map(item => validarDatosCurso(item));
                    
                    // Filtrar por categoría si se especificó
                    if (categoria) {
                        cursosData = todosCursos.filter(curso => {
                            const cursoCategoria = (curso.categoria || '').toLowerCase();
                            return cursoCategoria.includes(categoria);
                        });
                        console.log(`✅ Cursos de ${categoria}:`, cursosData.length);
                    } else {
                        cursosData = todosCursos;
                        console.log('✅ Todos los cursos cargados:', cursosData.length);
                    }
                    
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
                    nivel: 'Básico',
                    categoria: 'Fundamentación'
                },
                {
                    id: 2,
                    nombre: 'Curso de Escolta',
                    descripcion: 'Formación especializada en protección personal',
                    duracion: '160 horas',
                    precio: '$850.000',
                    modalidad: 'Presencial',
                    nivel: 'Avanzado',
                    categoria: 'Fundamentación'
                },
                {
                    id: 3,
                    nombre: 'Supervisor de Seguridad',
                    descripcion: 'Liderazgo y gestión de equipos de seguridad',
                    duracion: '80 horas',
                    precio: '$650.000',
                    modalidad: 'Híbrido',
                    nivel: 'Especialización',
                    categoria: 'Fundamentación'
                },
                {
                    id: 4,
                    nombre: 'Operador de Medios Tecnológicos',
                    descripcion: 'Monitoreo y control de sistemas de seguridad',
                    duracion: '100 horas',
                    precio: '$550.000',
                    modalidad: 'Presencial',
                    nivel: 'Especialización',
                    categoria: 'Fundamentación'
                },
                {
                    id: 5,
                    nombre: 'Reentrenamiento Vigilancia',
                    descripcion: 'Actualización para personal activo',
                    duracion: '30 horas',
                    precio: '$250.000',
                    modalidad: 'Virtual',
                    nivel: 'Actualización',
                    categoria: 'Reentrenamiento'
                },
                {
                    id: 6,
                    nombre: 'Manejador Canino',
                    descripcion: 'Entrenamiento con perros de seguridad',
                    duracion: '140 horas',
                    precio: '$950.000',
                    modalidad: 'Presencial',
                    nivel: 'Especialización',
                    categoria: 'Especialización'
                }
            ];
            
            // Filtrar por categoría si se especificó
            if (categoria) {
                cursosData = cursosData.filter(c => c.categoria.toLowerCase().includes(categoria));
            }
            
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
        nivel: item.Requisitos?.Value || item.Nivel || 'Básico',
        categoria: item.Categoria || item.Requisitos?.Value || item.Nivel || ''
    };
}

/**
 * Renderizar cursos en la página
 */
function renderizarCursos() {
    const container = document.getElementById('coursesContainer');
    
    if (cursosData.length === 0) {
        container.innerHTML = `
            <div style="text-align: center; padding: 60px 20px;">
                <p style="font-size: 18px; color: #64748b;">No hay cursos disponibles en esta categoría</p>
            </div>
        `;
        return;
    }
    
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
 * Subir archivo a biblioteca de documentos de SharePoint
 */
async function subirArchivoSharePoint(file, inscripcionId, tipoDocumento, formDigest) {
    try {
        if (!file) return null;
        
        const fileBuffer = await file.arrayBuffer();
        const fileName = `${inscripcionId}_${tipoDocumento}_${file.name}`;
        const folderUrl = `${CONFIG.sharepoint.bibliotecaDocumentos}`;

        const uploadResponse = await fetch(
            `${CONFIG.sharepoint.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/Files/add(url='${fileName}',overwrite=true)`,
            {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'X-RequestDigest': formDigest
                },
                body: fileBuffer
            }
        );

        if (uploadResponse.ok) {
            const uploadData = await uploadResponse.json();
            const fileUrl = uploadData.d.ServerRelativeUrl;
            const fullUrl = `${CONFIG.sharepoint.siteUrl}${fileUrl}`;
            console.log(`✅ Archivo ${tipoDocumento} subido:`, fullUrl);
            return fullUrl;
        }
        
        return null;
    } catch (error) {
        console.warn(`⚠️ No se pudo subir archivo ${tipoDocumento}:`, error);
        return null;
    }
}

/**
 * Guardar inscripción en SharePoint
 */
async function guardarEnSharePoint(formData, curso) {
    try {
        // Obtener Form Digest Token
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

        // Primero crear el item de inscripción (sin archivos)
        const itemData = {
            '__metadata': { 'type': 'SP.Data.InscripcionesListItem' },
            'Nombres': formData.nombres,
            'Apellidos': formData.apellidos,
            'Documento': formData.documento,
            'LugarExpedicion': formData.lugarExpedicion,
            'FechaNacimiento': formData.fechaNacimiento,
            'FechaExpedicion': formData.fechaExpedicion,
            'Correo': formData.correo,
            'Telefono': formData.telefono,
            'TipoEstudiante': formData.tipoEstudiante,
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
        const inscripcionId = data.d.Id;
        console.log('✅ Inscripción guardada en SharePoint:', inscripcionId);

        // Subir archivos y obtener URLs
        console.log('📤 Subiendo archivos a SharePoint...');
        const fotoUrl = await subirArchivoSharePoint(formData.fotoFile, inscripcionId, 'Foto', formDigest);
        const cedulaUrl = await subirArchivoSharePoint(formData.cedulaFile, inscripcionId, 'Cedula', formDigest);
        const cursoAnteriorUrl = await subirArchivoSharePoint(formData.cursoAnteriorFile, inscripcionId, 'CursoAnterior', formDigest);

        // Actualizar el item con las URLs de los archivos
        const updateData = {
            '__metadata': { 'type': 'SP.Data.InscripcionesListItem' }
        };

        if (fotoUrl) {
            updateData.Foto = {
                '__metadata': { 'type': 'SP.FieldUrlValue' },
                'Url': fotoUrl,
                'Description': 'Foto 3x4'
            };
        }

        if (cedulaUrl) {
            updateData.FotocopiaCedula = {
                '__metadata': { 'type': 'SP.FieldUrlValue' },
                'Url': cedulaUrl,
                'Description': 'Fotocopia Cédula'
            };
        }

        if (cursoAnteriorUrl) {
            updateData.UltimoCursoRealizado = {
                '__metadata': { 'type': 'SP.FieldUrlValue' },
                'Url': cursoAnteriorUrl,
                'Description': 'Certificado Curso Anterior'
            };
        }

        // Solo actualizar si hay archivos que agregar
        if (fotoUrl || cedulaUrl || cursoAnteriorUrl) {
            const updateResponse = await fetch(
                `${CONFIG.sharepoint.siteUrl}/_api/web/lists/getbytitle('${CONFIG.sharepoint.listaInscripciones}')/items(${inscripcionId})`,
                {
                    method: 'POST',
                    headers: {
                        'Accept': 'application/json;odata=verbose',
                        'Content-Type': 'application/json;odata=verbose',
                        'X-RequestDigest': formDigest,
                        'IF-MATCH': '*',
                        'X-HTTP-Method': 'MERGE'
                    },
                    body: JSON.stringify(updateData)
                }
            );

            if (updateResponse.ok) {
                console.log('✅ URLs de archivos guardadas en SharePoint');
            }
        }

        return inscripcionId;
    } catch (error) {
        console.warn('⚠️ No se pudo guardar en SharePoint:', error);
        // Retornar ID simulado en modo demo
        return Math.floor(Math.random() * 10000);
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
    doc.text(`${datosInscripcion.nombres} ${datosInscripcion.apellidos}`, 55, yPos);
    
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
    doc.text(curso.nombre, 55, yPos, { maxWidth: 130 });
    
    yPos += 12;
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
            tipoEstudiante: datosInscripcion.tipoEstudiante,
            nombres: datosInscripcion.nombres,
            apellidos: datosInscripcion.apellidos,
            documento: datosInscripcion.documento,
            lugarExpedicion: datosInscripcion.lugarExpedicion,
            fechaNacimiento: datosInscripcion.fechaNacimiento,
            fechaExpedicion: datosInscripcion.fechaExpedicion,
            telefono: datosInscripcion.telefono,
            curso: curso.nombre,
            precio: curso.precio,
            duracion: curso.duracion,
            modalidad: curso.modalidad,
            fotoBase64: datosInscripcion.fotoBase64 || '',
            fotoNombre: datosInscripcion.fotoNombre || '',
            cedulaBase64: datosInscripcion.cedulaBase64 || '',
            cedulaNombre: datosInscripcion.cedulaNombre || '',
            certificadoBase64: datosInscripcion.certificadoBase64 || '',
            certificadoNombre: datosInscripcion.certificadoNombre || '',
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
    document.getElementById('fotoFileName').textContent = 'Click para adjuntar foto';
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
    // Foto 3x4
    const fotoInput = document.getElementById('fotoFile');
    if (fotoInput) {
        fotoInput.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                document.getElementById('fotoFileName').textContent = e.target.files[0].name;
                console.log('📸 Foto seleccionada:', e.target.files[0].name);
            }
        });
    }
    
    // Cédula
    const cedulaInput = document.getElementById('cedulaFile');
    if (cedulaInput) {
        cedulaInput.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                document.getElementById('cedulaFileName').textContent = e.target.files[0].name;
                console.log('📄 Cédula seleccionada:', e.target.files[0].name);
            }
        });
    }
    
    // Curso anterior
    const cursoAnteriorInput = document.getElementById('cursoAnteriorFile');
    if (cursoAnteriorInput) {
        cursoAnteriorInput.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                document.getElementById('cursoAnteriorFileName').textContent = e.target.files[0].name;
                console.log('📄 Certificado seleccionado:', e.target.files[0].name);
            }
        });
    }
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
        tipoEstudiante: form.tipoEstudiante.value,
        nombres: form.nombres.value.trim(),
        apellidos: form.apellidos.value.trim(),
        documento: form.documento.value.trim(),
        lugarExpedicion: form.lugarExpedicion.value.trim(),
        fechaNacimiento: form.fechaNacimiento.value,
        fechaExpedicion: form.fechaExpedicion.value,
        correo: form.correo.value.trim(),
        telefono: form.telefono.value.trim(),
        fotoFile: form.foto.files[0] || null,
        cedulaFile: form.cedula.files[0],
        cursoAnteriorFile: form.cursoAnterior.files[0] || null
    };
    
    console.log('📋 Datos del formulario:', {
        tipoEstudiante: formData.tipoEstudiante,
        nombres: formData.nombres,
        apellidos: formData.apellidos,
        documento: formData.documento,
        curso: cursoSeleccionado.nombre
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
    
    if (formData.fotoFile && formData.fotoFile.size > 5 * 1024 * 1024) {
        mostrarNotificacion('La foto es muy grande. Máximo 5MB', 'error');
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
        console.log('💾 Paso 1/3: Guardando inscripción en SharePoint...');
        const inscripcionId = await guardarEnSharePoint(formData, cursoSeleccionado);
        console.log(`✅ Inscripción guardada con ID: ${inscripcionId}`);
        
        submitBtn.textContent = 'Procesando archivos...';

        // PASO 2: Convertir archivos a Base64
        console.log('📎 Paso 2/3: Convirtiendo archivos a Base64...');
        let fotoBase64 = '';
        let cedulaBase64 = '';
        let certificadoBase64 = '';

        if (formData.fotoFile) {
            fotoBase64 = await convertirArchivoABase64(formData.fotoFile);
            console.log('✅ Foto convertida a Base64');
        }

        if (formData.cedulaFile) {
            cedulaBase64 = await convertirArchivoABase64(formData.cedulaFile);
            console.log('✅ Cédula convertida a Base64');
        }

        if (formData.cursoAnteriorFile) {
            certificadoBase64 = await convertirArchivoABase64(formData.cursoAnteriorFile);
            console.log('✅ Certificado convertido a Base64');
        }

        submitBtn.textContent = 'Enviando correo de confirmación...';

        // PASO 3: Enviar correo con PDF y archivos adjuntos
        console.log('📧 Paso 3/3: Enviando correo...');

        // Crear datos completos con archivos
        const datosCompletos = {
            ...formData,
            fotoBase64: fotoBase64,
            fotoNombre: formData.fotoFile ? formData.fotoFile.name : '',
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
window.seleccionarCategoria = seleccionarCategoria;
window.volverACategorias = volverACategorias;
window.abrirFormulario = abrirFormulario;
window.cerrarFormulario = cerrarFormulario;
window.enviarInscripcion = enviarInscripcion;

console.log('✅ ADEVIP Functions cargadas correctamente');
console.log('🔧 Configuración:', {
    sharepoint: CONFIG.sharepoint.siteUrl,
    flowConfigured: CONFIG.flowUrl.includes('powerplatform') ? 'SÍ' : 'NO'
});
