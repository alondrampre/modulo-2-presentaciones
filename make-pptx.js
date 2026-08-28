const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

// 1. Leer slides.js
const slidesCode = fs.readFileSync(path.join(__dirname, 'slides.js'), 'utf8');

// 2. Ejecutar slides.js en un entorno de sandbox
const sandbox = {
  slidesData: {},
  missingSyllabusData: {}
};

const cleanCode = slidesCode
  .replace(/const slidesData/g, 'slidesData')
  .replace(/const missingSyllabusData/g, 'missingSyllabusData')
  .replace(/window\.addEventListener.*/g, '')
  .replace(/document\.addEventListener.*/g, '')
  .replace(/cleanDatabaseMedia\(\);/g, '')
  .replace(/assignCardIconsAndStyles\(\);/g, '');

try {
  const fn = new Function('slidesData', 'missingSyllabusData', cleanCode + '\nreturn {slidesData, missingSyllabusData};');
  const result = fn({}, {});
  sandbox.slidesData = result.slidesData;
  sandbox.missingSyllabusData = result.missingSyllabusData;
} catch (err) {
  console.error("Error al evaluar slides.js:", err);
}

// 3. Crear la presentación PowerPoint
let pptx = new pptxgen();
pptx.layout = 'LAYOUT_16x9';

const colors = {
  bg: '151F30',      // Navy oscuro (#151f30)
  title: 'FFBB27',   // Amarillo (#ffbb27)
  text: 'FFFFFF',    // Blanco (#ffffff)
  coral: 'FF7448',   // Coral (#ff7448)
  cardBg: '1E293B',  // Ligeramente más claro (#1e293b)
  teal: '058080'     // Teal (#058080)
};

// Función para parsear HTML simple y listas a fragmentos de texto multi-estilo de pptxgen
function parseHtmlToPptxText(htmlContent, defaultColor = 'FFFFFF') {
  const fragments = [];
  
  let temp = htmlContent
    .replace(/<p>/g, '')
    .replace(/<\/p>/g, '\n\n')
    .replace(/<br\s*\/?>/g, '\n');
    
  const listRegex = /<li[^>]*>(.*?)<\/li>/gi;
  let listMatch;
  let hasLists = false;
  let listItems = [];
  
  while ((listMatch = listRegex.exec(temp)) !== null) {
    hasLists = true;
    listItems.push(listMatch[1]);
  }
  
  if (hasLists) {
    listItems.forEach((item, index) => {
      const bulletColor = (index % 2 === 0) ? colors.coral : colors.title;
      const highlightColor = (index % 2 === 0) ? colors.coral : colors.title;
      
      fragments.push({ text: "→  ", options: { color: bulletColor, bold: true, fontFace: 'Arial' } });
      
      const strongRegex = /<strong>(.*?)<\/strong>|<b>(.*?)<\/b>/i;
      const strongMatch = item.match(strongRegex);
      
      if (strongMatch) {
        const strongText = strongMatch[1] || strongMatch[2];
        const restText = item.replace(strongRegex, '').replace(/<[^>]*>/g, '');
        
        fragments.push({ text: strongText + ": ", options: { color: highlightColor, bold: true, fontFace: 'Century Gothic' } });
        fragments.push({ text: restText + '\n', options: { color: defaultColor, fontFace: 'Century Gothic' } });
      } else {
        const cleanItem = item.replace(/<[^>]*>/g, '');
        fragments.push({ text: cleanItem + '\n', options: { color: defaultColor, fontFace: 'Century Gothic' } });
      }
    });
  } else {
    // Si contiene negritas en texto libre, las separamos
    const parts = temp.split(/(<strong>.*?<\/strong>)/gi);
    parts.forEach(part => {
      if (part.startsWith('<strong>')) {
        const cleanPart = part.replace(/<[^>]*>/g, '');
        fragments.push({ text: cleanPart, options: { color: colors.title, bold: true, fontFace: 'Century Gothic' } });
      } else {
        const cleanPart = part.replace(/<[^>]*>/g, '');
        fragments.push({ text: cleanPart, options: { color: defaultColor, fontFace: 'Century Gothic' } });
      }
    });
  }
  
  return fragments;
}

// Recorrer la base de datos de diapositivas en orden de semana y día
const weekKeys = ['semana-5', 'semana-6', 'semana-7', 'semana-8'];
const dayKeys = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes'];

let classCounter = 21;
const defaultIcons = ["💡", "⚙️", "🎨", "🔬", "🌱", "🎯", "🔍", "💼", "💰", "👥", "📢", "🏺", "⏳", "📐", "🎬", "🔑"];

weekKeys.forEach(weekKey => {
  const week = sandbox.slidesData[weekKey];
  if (!week) return;
  
  dayKeys.forEach(dayKey => {
    const day = week.days[dayKey];
    if (!day || !day.slides || day.slides.length === 0) return;
    
    day.slides.forEach((slide, slideIndex) => {
      let pptxSlide = pptx.addSlide();
      pptxSlide.background = { color: colors.bg };
      
      // Encabezado con número de diapositiva y barra de color
      pptxSlide.addText((slideIndex + 1).toString().padStart(2, '0'), {
        x: 0.6,
        y: 0.3,
        w: 1.0,
        h: 0.4,
        fontSize: 24,
        fontFace: 'Georgia',
        color: colors.coral,
        bold: true
      });
      
      pptxSlide.addShape(pptx.shapes.RECTANGLE, {
        x: 0.6,
        y: 0.7,
        w: 0.8,
        h: 0.05,
        fill: { color: colors.coral }
      });
      
      // Texto del pie de página
      const footerText = `MÓDULO 2: CONSTRUCCIÓN DE MARCA 360° • ${dayKey.toUpperCase()}: ${day.title ? day.title.split(': ')[1] : ''}`;
      pptxSlide.addText(footerText, {
        x: 0.6,
        y: 7.0,
        w: 8.0,
        h: 0.3,
        fontSize: 9,
        fontFace: 'Arial',
        color: '718096',
        opacity: 0.7
      });
      
      pptxSlide.addText('Lexpin', {
        x: 11.5,
        y: 7.0,
        w: 1.5,
        h: 0.3,
        fontSize: 10,
        fontFace: 'Georgia',
        color: colors.coral,
        bold: true,
        align: 'right',
        opacity: 0.8
      });
      
      if (slide.layout === 'title') {
        const isQuestion = slide.title.includes('?') || slide.title.toLowerCase().includes('actividad') || slide.title.toLowerCase().includes('práctica');
        
        let titleX = 1.0;
        let titleY = 2.7;
        let subX = 2.0;
        let subY = 5.1;
        
        if (slide.positions) {
          if (slide.positions.title) {
            titleX = (slide.positions.title.left / 100) * 13.33;
            titleY = (slide.positions.title.top / 100) * 7.5;
          }
          if (slide.positions.content) {
            subX = (slide.positions.content.left / 100) * 13.33;
            subY = (slide.positions.content.top / 100) * 7.5;
          }
        }
        
        let pillY = Math.max(0.8, titleY - 0.9);
        
        if (isQuestion) {
          pptxSlide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 3.5,
            y: pillY,
            w: 5.6,
            h: 0.6,
            fill: { color: '1A202C' },
            line: { color: '2D3748', width: 1.5 },
            rectRadius: 0.5
          });
          pptxSlide.addText('@lexpinonline | Lexpin Cursos Online', {
            x: 3.5,
            y: pillY,
            w: 5.6,
            h: 0.6,
            fontSize: 12,
            fontFace: 'Century Gothic',
            color: colors.title,
            align: 'center',
            valign: 'middle',
            bold: true
          });
        } else {
          pptxSlide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 3.5,
            y: pillY,
            w: 5.6,
            h: 0.6,
            fill: { color: '1A202C' },
            line: { color: '2D3748', width: 1.5 },
            rectRadius: 0.5
          });
          pptxSlide.addText(`MÓDULO 2 • CLASE ${classCounter}`, {
            x: 3.5,
            y: pillY,
            w: 5.6,
            h: 0.6,
            fontSize: 11,
            fontFace: 'Century Gothic',
            color: colors.title,
            align: 'center',
            valign: 'middle',
            bold: true
          });
        }
        
        // Título central
        pptxSlide.addText(slide.title, {
          x: titleX,
          y: titleY,
          w: 11.33,
          h: 2.2,
          fontSize: 42,
          fontFace: 'Georgia',
          color: colors.text,
          align: 'center',
          valign: 'middle',
          bold: true
        });
        
        // Subtítulo
        if (slide.content) {
          const cleanText = slide.content.replace(/<[^>]*>/g, '').trim();
          pptxSlide.addText(cleanText, {
            x: subX,
            y: subY,
            w: 9.33,
            h: 1.2,
            fontSize: 16,
            fontFace: 'Century Gothic',
            color: 'A0AEC0',
            align: 'center'
          });
        }
        
      } else if (slide.layout === 'quote') {
        let quoteX = 1.5;
        let quoteY = 2.4;
        let authorX = 2.0;
        let authorY = 5.2;
        
        if (slide.positions) {
          if (slide.positions.quote) {
            quoteX = (slide.positions.quote.left / 100) * 13.33;
            quoteY = (slide.positions.quote.top / 100) * 7.5;
          }
          if (slide.positions.author) {
            authorX = (slide.positions.author.left / 100) * 13.33;
            authorY = (slide.positions.author.top / 100) * 7.5;
          }
        }
        
        // Comillas gigantes
        pptxSlide.addText('“', {
          x: Math.max(0.4, quoteX - 0.7),
          y: Math.max(0.8, quoteY - 0.9),
          w: 1.0,
          h: 1.0,
          fontSize: 80,
          fontFace: 'Georgia',
          color: colors.coral,
          bold: true
        });
        
        const cleanQuote = slide.content.replace(/<[^>]*>/g, '').trim();
        pptxSlide.addText(cleanQuote, {
          x: quoteX,
          y: quoteY,
          w: 10.33,
          h: 2.5,
          fontSize: 24,
          fontFace: 'Georgia',
          color: colors.text,
          align: 'center',
          italic: true
        });
        
        pptxSlide.addText('”', {
          x: Math.min(12.0, quoteX + 10.0),
          y: Math.min(6.5, quoteY + 2.0),
          w: 1.0,
          h: 1.0,
          fontSize: 80,
          fontFace: 'Georgia',
          color: colors.coral,
          bold: true
        });
        
        pptxSlide.addText(`— ${slide.title}`, {
          x: authorX,
          y: authorY,
          w: 9.33,
          h: 0.5,
          fontSize: 18,
          fontFace: 'Century Gothic',
          color: colors.title,
          align: 'center',
          bold: true
        });
        
      } else if (slide.layout === 'split') {
        let titleX = 0.8;
        let titleY = 1.0;
        let textX = 0.8;
        let textY = 2.0;
        let mediaX = 6.8;
        let mediaY = 1.6;
        
        if (slide.positions) {
          if (slide.positions.title) {
            titleX = (slide.positions.title.left / 100) * 13.33;
            titleY = (slide.positions.title.top / 100) * 7.5;
          }
          if (slide.positions.content) {
            textX = (slide.positions.content.left / 100) * 13.33;
            textY = (slide.positions.content.top / 100) * 7.5;
          }
          if (slide.positions.media) {
            mediaX = (slide.positions.media.left / 100) * 13.33;
            mediaY = (slide.positions.media.top / 100) * 7.5;
          }
        }
        
        // Título de la diapositiva
        pptxSlide.addText(slide.title, {
          x: titleX,
          y: titleY,
          w: 5.5,
          h: 0.8,
          fontSize: 28,
          fontFace: 'Georgia',
          color: colors.title,
          bold: true
        });
        
        // Columna de texto izquierda con multi-estilo
        const textFragments = parseHtmlToPptxText(slide.content);
        pptxSlide.addText(textFragments, {
          x: textX,
          y: textY,
          w: 5.5,
          h: 4.5,
          fontSize: 14,
          lineSpacing: 22
        });
        
        // Columna de medios derecha (Tarjeta premium)
        pptxSlide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: mediaX,
          y: mediaY,
          w: 5.5,
          h: 4.8,
          fill: { color: colors.cardBg },
          line: { color: '2D3748', width: 1.5 },
          rectRadius: 0.05
        });
        
        let mediaLabel = '[ Diagrama Conceptual ]';
        if (slide.media && !slide.media.startsWith('<svg')) {
          mediaLabel = `[ Referencia Visual: ${path.basename(slide.media)} ]`;
        }
        
        pptxSlide.addText(mediaLabel, {
          x: mediaX + 0.2,
          y: mediaY + 2.0,
          w: 5.1,
          h: 0.8,
          fontSize: 14,
          fontFace: 'Century Gothic',
          color: colors.title,
          align: 'center',
          bold: true
        });
        
      } else if (slide.layout === 'grid' && slide.cards) {
        let titleX = 0.8;
        let titleY = 1.0;
        let contentX = 0.8;
        let contentY = 1.7;
        let startY = 1.8;
        
        if (slide.positions) {
          if (slide.positions.title) {
            titleX = (slide.positions.title.left / 100) * 13.33;
            titleY = (slide.positions.title.top / 100) * 7.5;
          }
          if (slide.positions.content) {
            contentX = (slide.positions.content.left / 100) * 13.33;
            contentY = (slide.positions.content.top / 100) * 7.5;
            startY = contentY + 0.4;
          }
        }
        
        // Título
        pptxSlide.addText(slide.title, {
          x: titleX,
          y: titleY,
          w: 11.7,
          h: 0.8,
          fontSize: 28,
          fontFace: 'Georgia',
          color: colors.title,
          bold: true
        });
        
        if (slide.content && slide.content !== "<h2>Verificación antes de exportar:</h2>" && slide.content !== "<h2>Antes de la entrega final del lunes, comprueben:</h2>") {
          const cleanGridText = slide.content.replace(/<[^>]*>/g, '').trim();
          pptxSlide.addText(cleanGridText, {
            x: contentX,
            y: contentY,
            w: 11.7,
            h: 0.4,
            fontSize: 14,
            fontFace: 'Century Gothic',
            color: 'A0AEC0'
          });
          if (!slide.positions || !slide.positions.content) {
            startY = 2.2;
          }
        }
        
        // Dibujar tarjetas
        const numCards = slide.cards.length;
        const availableWidth = 11.7;
        const gap = 0.3;
        const cardW = (availableWidth - (numCards - 1) * gap) / numCards;
        const cardH = 4.2;
        
        slide.cards.forEach((card, i) => {
          let cardX = 0.8 + i * (cardW + gap);
          let cardY = startY;
          
          if (slide.positions && slide.positions.cards && slide.positions.cards[i]) {
            cardX = (slide.positions.cards[i].left / 100) * 13.33;
            cardY = (slide.positions.cards[i].top / 100) * 7.5;
          }
          
          // Fondo de la tarjeta (Estilo tarjeta de la web)
          pptxSlide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: cardX,
            y: cardY,
            w: cardW,
            h: cardH,
            fill: { color: colors.cardBg },
            line: { color: '2D3748', width: 1.5 },
            rectRadius: 0.05
          });
          
          let textY = cardY + 0.3;
          const cardIcon = card.icon || defaultIcons[i % defaultIcons.length];
          
          // Badge circular azul-teal para el icono (Estilo Modelo 4)
          pptxSlide.addShape(pptx.shapes.OVAL, {
            x: cardX + 0.3,
            y: cardY + 0.25,
            w: 0.65,
            h: 0.65,
            fill: { color: '0E2A36' },
            line: { color: colors.teal, width: 1.5 }
          });
          
          pptxSlide.addText(cardIcon, {
            x: cardX + 0.3,
            y: cardY + 0.25,
            w: 0.65,
            h: 0.65,
            fontSize: 18,
            align: 'center',
            valign: 'middle'
          });
          textY = cardY + 1.0;
          
          // Título de la tarjeta
          pptxSlide.addText(card.title, {
            x: cardX + 0.3,
            y: textY,
            w: cardW - 0.6,
            h: 0.5,
            fontSize: 14,
            fontFace: 'Georgia',
            color: colors.text,
            bold: true
          });
          
          // Descripción de la tarjeta con soporte para listas (Estilo Modelo 5)
          const descFragments = parseHtmlToPptxText(card.desc, 'CBD5E0');
          pptxSlide.addText(descFragments, {
            x: cardX + 0.3,
            y: textY + 0.6,
            w: cardW - 0.6,
            h: cardH - (textY - cardY) - 0.8,
            fontSize: 11,
            lineSpacing: 16
          });
        });
      }
    });
    
    classCounter++;
  });
});

// Guardar presentación directamente en el Escritorio del usuario (alond)
const desktopPptxFile = 'C:/Users/alond/OneDrive/Escritorio/presentacion-modulo-2.pptx';
pptx.writeFile({ fileName: desktopPptxFile })
  .then(fileName => {
    console.log(`PowerPoint creado exitosamente en el Escritorio: ${fileName}`);
    
    // Copiar también el archivo ZIP al Escritorio de forma nativa en Node
    try {
      const srcZip = 'C:/Users/alond/.gemini/antigravity/brain/9621502e-b2fa-4f18-908b-471b6d8f727d/modulo-2-presentaciones.zip';
      const destZip = 'C:/Users/alond/OneDrive/Escritorio/modulo-2-presentaciones.zip';
      if (fs.existsSync(srcZip)) {
        fs.copyFileSync(srcZip, destZip);
        console.log(`ZIP copiado exitosamente al Escritorio: ${destZip}`);
      }
    } catch (zipErr) {
      console.error('Error al copiar ZIP al Escritorio:', zipErr);
    }
  })
  .catch(err => {
    console.error('Error al guardar PowerPoint:', err);
  });
