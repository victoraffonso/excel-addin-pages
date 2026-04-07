/**
 * PowerPoint Operations — maps MCP tool names to Office.js PowerPoint API calls.
 * Each function receives a PowerPoint.RequestContext and tool arguments.
 * Operations are queued and executed in a single context.sync() by the caller.
 *
 * Note: Office.js PowerPoint API is more limited than Excel. Some operations
 * (charts, tables, images, speaker notes) are not fully supported and will
 * throw clear errors directing users to the python-pptx MCP fallback.
 */

/**
 * Main dispatcher — routes tool name to the correct operation function.
 */
export async function executeOperation(context, tool, args) {
  const handler = OPERATIONS[tool];
  if (!handler) {
    throw new Error(`Unknown operation: ${tool}`);
  }
  return await handler(context, args);
}

// Helper: convert inches to PowerPoint points (1 inch = 72 points)
function inchesToPoints(inches) {
  return inches * 72;
}

// Helper: get slide by 0-based index
function getSlide(context, slideIndex) {
  const slides = context.presentation.slides;
  slides.load('items');
  return { slides, index: slideIndex };
}

const OPERATIONS = {
  // ─── Presentation Info ───
  async get_presentation_info(context, args) {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    const slideInfos = [];
    for (let i = 0; i < slides.items.length; i++) {
      const slide = slides.items[i];
      slide.shapes.load('items');
      await context.sync();

      let title = '';
      for (const shape of slide.shapes.items) {
        if (shape.type === 'Placeholder' || shape.name?.toLowerCase().includes('title')) {
          shape.textFrame.load('textRange');
          try {
            await context.sync();
            shape.textFrame.textRange.load('text');
            await context.sync();
            title = shape.textFrame.textRange.text || '';
            if (title) break;
          } catch {
            // Shape may not have a text frame
          }
        }
      }

      // If no title found, try first shape with text
      if (!title && slide.shapes.items.length > 0) {
        try {
          const firstShape = slide.shapes.items[0];
          firstShape.textFrame.load('textRange');
          await context.sync();
          firstShape.textFrame.textRange.load('text');
          await context.sync();
          title = firstShape.textFrame.textRange.text || '';
        } catch {
          // No text in first shape
        }
      }

      slideInfos.push({ index: i, title });
    }

    return {
      slide_count: slides.items.length,
      slides: slideInfos,
    };
  },

  async list_layouts(context, args) {
    // Office.js PowerPoint API doesn't fully expose slide layouts for enumeration.
    // Return a standard set of layout descriptions.
    return [
      { index: 0, name: 'Title Slide' },
      { index: 1, name: 'Title and Content' },
      { index: 2, name: 'Section Header' },
      { index: 3, name: 'Two Content' },
      { index: 4, name: 'Comparison' },
      { index: 5, name: 'Title Only' },
      { index: 6, name: 'Blank' },
      { index: 7, name: 'Content with Caption' },
      { index: 8, name: 'Picture with Caption' },
    ];
  },

  // ─── Slide Operations ───
  async add_slide(context, args) {
    const slides = context.presentation.slides;
    slides.add();
    await context.sync();

    // Reload to get count
    slides.load('items');
    await context.sync();
    return `Added slide ${slides.items.length} — visible in PowerPoint`;
  },

  async delete_slide(context, args) {
    const slideIndex = args.slide_index;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slideIndex < 0 || slideIndex >= slides.items.length) {
      throw new Error(`Slide index ${slideIndex} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slideIndex];
    slide.delete();
    return `Deleted slide ${slideIndex}`;
  },

  async duplicate_slide(context, args) {
    // Office.js PowerPoint doesn't have a direct duplicate method.
    // We need to use the slides.add() and then copy content approach,
    // but content copying is limited. Use insertSlidesFromBase64 as an alternative.
    throw new Error('Slide duplication is not directly supported via Office.js add-in. Use the python-pptx MCP for this operation: mcp__powerpoint__duplicate_slide');
  },

  // ─── Slide Info ───
  async get_slide_info(context, args) {
    const slideIndex = args.slide_index;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slideIndex < 0 || slideIndex >= slides.items.length) {
      throw new Error(`Slide index ${slideIndex} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slideIndex];
    slide.shapes.load('items');
    await context.sync();

    const shapes = [];
    for (let i = 0; i < slide.shapes.items.length; i++) {
      const shape = slide.shapes.items[i];
      shape.load('name,type,left,top,width,height');
      await context.sync();

      let text = '';
      let hasText = false;
      try {
        shape.textFrame.load('textRange,hasText');
        await context.sync();
        hasText = shape.textFrame.hasText;
        if (hasText) {
          shape.textFrame.textRange.load('text');
          await context.sync();
          text = shape.textFrame.textRange.text || '';
        }
      } catch {
        // Shape doesn't have a text frame (e.g., images, charts)
      }

      shapes.push({
        index: i + 1, // 1-based to match python-pptx MCP convention
        name: shape.name,
        type: shape.type,
        has_text: hasText,
        text: text,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      });
    }

    return {
      slide_index: slideIndex,
      shape_count: slide.shapes.items.length,
      shapes,
    };
  },

  // ─── Text Operations ───
  async set_slide_text(context, args) {
    const { slide_index, shape_index, text } = args;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slide_index < 0 || slide_index >= slides.items.length) {
      throw new Error(`Slide index ${slide_index} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slide_index];
    slide.shapes.load('items');
    await context.sync();

    // shape_index is 1-based per python-pptx MCP convention
    const shapeIdx = shape_index - 1;
    if (shapeIdx < 0 || shapeIdx >= slide.shapes.items.length) {
      throw new Error(`Shape index ${shape_index} out of range (1-${slide.shapes.items.length})`);
    }

    const shape = slide.shapes.items[shapeIdx];
    shape.textFrame.textRange.load('text');
    await context.sync();
    shape.textFrame.textRange.text = text;

    return `Text set on slide ${slide_index}, shape ${shape_index}`;
  },

  async set_title(context, args) {
    // Title is shape 1 (index 0) by convention
    return await OPERATIONS.set_slide_text(context, {
      ...args,
      shape_index: 1,
    });
  },

  async format_text(context, args) {
    const { slide_index, shape_index, font_name, font_size, bold, italic, color } = args;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slide_index < 0 || slide_index >= slides.items.length) {
      throw new Error(`Slide index ${slide_index} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slide_index];
    slide.shapes.load('items');
    await context.sync();

    const shapeIdx = shape_index - 1;
    if (shapeIdx < 0 || shapeIdx >= slide.shapes.items.length) {
      throw new Error(`Shape index ${shape_index} out of range (1-${slide.shapes.items.length})`);
    }

    const shape = slide.shapes.items[shapeIdx];
    const font = shape.textFrame.textRange.font;

    if (font_name != null) font.name = font_name;
    if (font_size != null) font.size = font_size;
    if (bold != null) font.bold = bold;
    if (italic != null) font.italic = italic;
    if (color != null) font.color = '#' + color.replace('#', '');

    return `Formatted shape ${shape_index} on slide ${slide_index}`;
  },

  async add_text_box(context, args) {
    const { slide_index, text, left = 1.0, top = 1.0, width = 5.0, height = 1.0 } = args;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slide_index < 0 || slide_index >= slides.items.length) {
      throw new Error(`Slide index ${slide_index} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slide_index];

    // addTextBox takes points (1 inch = 72 points)
    const options = {
      left: inchesToPoints(left),
      top: inchesToPoints(top),
      width: inchesToPoints(width),
      height: inchesToPoints(height),
    };

    const textBox = slide.shapes.addTextBox(text, options);
    textBox.load('name');
    await context.sync();

    return `Text box added on slide ${slide_index}: "${text.substring(0, 40)}${text.length > 40 ? '...' : ''}"`;
  },

  // ─── Read Operations ───
  async read_all_text(context, args) {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    const result = [];
    for (let i = 0; i < slides.items.length; i++) {
      const slide = slides.items[i];
      slide.shapes.load('items');
      await context.sync();

      const texts = [];
      let title = '';

      for (let s = 0; s < slide.shapes.items.length; s++) {
        const shape = slide.shapes.items[s];
        try {
          shape.textFrame.load('hasText');
          await context.sync();
          if (shape.textFrame.hasText) {
            shape.textFrame.textRange.load('text');
            await context.sync();
            const txt = shape.textFrame.textRange.text || '';
            if (txt) {
              if (s === 0) title = txt;
              texts.push(txt);
            }
          }
        } catch {
          // Shape without text frame (images, etc.)
        }
      }

      result.push({ slide_index: i, title, texts });
    }
    return result;
  },

  // ─── Speaker Notes ───
  async set_speaker_notes(context, args) {
    // Office.js PowerPoint API does not support speaker notes manipulation.
    throw new Error('Speaker notes are not supported via Office.js add-in. Use the python-pptx MCP for this operation: mcp__powerpoint__set_speaker_notes');
  },

  async get_speaker_notes(context, args) {
    // Office.js PowerPoint API does not support speaker notes read.
    throw new Error('Speaker notes are not supported via Office.js add-in. Use the python-pptx MCP for this operation: mcp__powerpoint__get_speaker_notes');
  },

  // ─── Charts ───
  async add_chart(context, args) {
    throw new Error('Chart creation is not supported via Office.js add-in. Use the python-pptx MCP for this operation: mcp__powerpoint__add_chart');
  },

  // ─── Tables ───
  async add_table(context, args) {
    throw new Error('Table creation is not supported via Office.js add-in. Use the python-pptx MCP for this operation: mcp__powerpoint__add_table');
  },

  // ─── Images ───
  async add_image(context, args) {
    // Office.js supports addImage via base64, but we'd need the image data.
    // The MCP tool passes a file path, not base64. Delegate to python-pptx.
    throw new Error('Image insertion from file path is not supported via Office.js add-in. Use the python-pptx MCP for this operation: mcp__powerpoint__add_image');
  },

  // ─── Save & Navigate ───
  async save_presentation(context, args) {
    // Office.js doesn't have a direct save() method for PowerPoint.
    // The document auto-saves, or we use Office.context.document.saveAsync.
    return new Promise((resolve, reject) => {
      Office.context.document.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(`Saved: ${args.file_path || 'active presentation'}`);
        } else {
          reject(new Error(`Save failed: ${result.error?.message || 'unknown error'}`));
        }
      });
    });
  },

  async go_to_slide(context, args) {
    const slideIndex = args.slide_index;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slideIndex < 0 || slideIndex >= slides.items.length) {
      throw new Error(`Slide index ${slideIndex} out of range (0-${slides.items.length - 1})`);
    }

    // Use Office.context.document.goToByIdAsync for navigation
    // PowerPoint slides are 1-based in the goTo API
    return new Promise((resolve, reject) => {
      Office.context.document.goToByIdAsync(slideIndex + 1, Office.GoToType.Slide, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(`Navigated to slide ${slideIndex}`);
        } else {
          reject(new Error(`Navigation failed: ${result.error?.message || 'unknown error'}`));
        }
      });
    });
  },

  // ─── Shape Operations ───
  async delete_shape(context, args) {
    const { slide_index, shape_index } = args;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slide_index < 0 || slide_index >= slides.items.length) {
      throw new Error(`Slide index ${slide_index} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slide_index];
    slide.shapes.load('items');
    await context.sync();

    // shape_index is 1-based per convention
    const shapeIdx = shape_index - 1;
    if (shapeIdx < 0 || shapeIdx >= slide.shapes.items.length) {
      throw new Error(`Shape index ${shape_index} out of range (1-${slide.shapes.items.length})`);
    }

    const shape = slide.shapes.items[shapeIdx];
    shape.delete();
    await context.sync();

    return `Deleted shape ${shape_index} from slide ${slide_index}`;
  },

  async add_shape(context, args) {
    const { slide_index, shape_type = 'rectangle', left = 2.0, top = 2.0, width = 3.0, height = 2.0, text = '' } = args;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slide_index < 0 || slide_index >= slides.items.length) {
      throw new Error(`Slide index ${slide_index} out of range (0-${slides.items.length - 1})`);
    }

    // Map friendly names to PowerPoint GeometricShapeType enum values
    const shapeTypeMap = {
      'rectangle': 'Rectangle',
      'oval': 'Ellipse',
      'ellipse': 'Ellipse',
      'triangle': 'Triangle',
      'diamond': 'Diamond',
      'roundedrectangle': 'RoundedRectangle',
      'rounded_rectangle': 'RoundedRectangle',
      'roundedRect': 'RoundedRectangle',
    };

    const mappedType = shapeTypeMap[shape_type.toLowerCase()] || shape_type;

    const slide = slides.items[slide_index];
    const options = {
      left: inchesToPoints(left),
      top: inchesToPoints(top),
      width: inchesToPoints(width),
      height: inchesToPoints(height),
    };

    const newShape = slide.shapes.addGeometricShape(mappedType, options);
    newShape.load('name');
    await context.sync();

    // Set text if provided
    if (text) {
      try {
        newShape.textFrame.textRange.text = text;
        await context.sync();
      } catch {
        // Some geometric shapes may not support text frames
      }
    }

    return `Added ${shape_type} shape on slide ${slide_index} (name: ${newShape.name})`;
  },

  async list_shapes(context, args) {
    const slideIndex = args.slide_index;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slideIndex < 0 || slideIndex >= slides.items.length) {
      throw new Error(`Slide index ${slideIndex} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slideIndex];
    slide.shapes.load('items');
    await context.sync();

    const shapes = [];
    for (let i = 0; i < slide.shapes.items.length; i++) {
      const shape = slide.shapes.items[i];
      shape.load('id,name,type,left,top,width,height');
      await context.sync();

      let text = '';
      let hasText = false;
      try {
        shape.textFrame.load('hasText');
        await context.sync();
        hasText = shape.textFrame.hasText;
        if (hasText) {
          shape.textFrame.textRange.load('text');
          await context.sync();
          const fullText = shape.textFrame.textRange.text || '';
          // Include full text up to 200 chars
          text = fullText.length > 200 ? fullText.substring(0, 200) + '...' : fullText;
        }
      } catch {
        // Shape without text frame
      }

      shapes.push({
        index: i + 1, // 1-based
        id: shape.id,
        name: shape.name,
        type: shape.type,
        left_inches: Math.round((shape.left / 72) * 100) / 100,
        top_inches: Math.round((shape.top / 72) * 100) / 100,
        width_inches: Math.round((shape.width / 72) * 100) / 100,
        height_inches: Math.round((shape.height / 72) * 100) / 100,
        has_text: hasText,
        text_preview: text,
      });
    }

    return {
      slide_index: slideIndex,
      shape_count: slide.shapes.items.length,
      shapes,
    };
  },

  async set_shape_position(context, args) {
    const { slide_index, shape_index, left, top, width, height } = args;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slide_index < 0 || slide_index >= slides.items.length) {
      throw new Error(`Slide index ${slide_index} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slide_index];
    slide.shapes.load('items');
    await context.sync();

    const shapeIdx = shape_index - 1;
    if (shapeIdx < 0 || shapeIdx >= slide.shapes.items.length) {
      throw new Error(`Shape index ${shape_index} out of range (1-${slide.shapes.items.length})`);
    }

    const shape = slide.shapes.items[shapeIdx];
    if (left != null) shape.left = inchesToPoints(left);
    if (top != null) shape.top = inchesToPoints(top);
    if (width != null) shape.width = inchesToPoints(width);
    if (height != null) shape.height = inchesToPoints(height);
    await context.sync();

    return `Repositioned shape ${shape_index} on slide ${slide_index}`;
  },

  async format_shape(context, args) {
    const { slide_index, shape_index, fill_color, line_color, line_width } = args;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slide_index < 0 || slide_index >= slides.items.length) {
      throw new Error(`Slide index ${slide_index} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slide_index];
    slide.shapes.load('items');
    await context.sync();

    const shapeIdx = shape_index - 1;
    if (shapeIdx < 0 || shapeIdx >= slide.shapes.items.length) {
      throw new Error(`Shape index ${shape_index} out of range (1-${slide.shapes.items.length})`);
    }

    const shape = slide.shapes.items[shapeIdx];
    const changes = [];

    if (fill_color != null) {
      shape.fill.setSolidColor('#' + fill_color.replace('#', ''));
      changes.push('fill');
    }
    if (line_color != null) {
      shape.lineFormat.color = '#' + line_color.replace('#', '');
      changes.push('line color');
    }
    if (line_width != null) {
      shape.lineFormat.weight = line_width;
      changes.push('line width');
    }

    await context.sync();
    return `Formatted shape ${shape_index} on slide ${slide_index} (${changes.join(', ')})`;
  },

  // ─── Slide Reorder ───
  async move_slide(context, args) {
    const { from_index, to_index } = args;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (from_index < 0 || from_index >= slides.items.length) {
      throw new Error(`From index ${from_index} out of range (0-${slides.items.length - 1})`);
    }
    if (to_index < 0 || to_index >= slides.items.length) {
      throw new Error(`To index ${to_index} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[from_index];
    // moveTo is 1-based position in PowerPointApi 1.5+
    try {
      slide.moveTo(to_index + 1);
      await context.sync();
    } catch (e) {
      throw new Error(`Slide moveTo failed — requires PowerPointApi 1.5+. Error: ${e.message}`);
    }

    return `Moved slide from position ${from_index} to ${to_index}`;
  },

  // ─── Search & Replace ───
  async search_replace_text(context, args) {
    const { search, replace } = args;
    if (!search) throw new Error('search parameter is required');

    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    let totalReplacements = 0;
    const slidesAffected = [];

    for (let i = 0; i < slides.items.length; i++) {
      const slide = slides.items[i];
      slide.shapes.load('items');
      await context.sync();

      let slideHits = 0;
      for (let s = 0; s < slide.shapes.items.length; s++) {
        const shape = slide.shapes.items[s];
        try {
          shape.textFrame.load('hasText');
          await context.sync();
          if (!shape.textFrame.hasText) continue;

          shape.textFrame.textRange.load('text');
          await context.sync();

          const originalText = shape.textFrame.textRange.text || '';
          if (originalText.includes(search)) {
            const newText = originalText.split(search).join(replace || '');
            shape.textFrame.textRange.text = newText;
            const hits = (originalText.split(search).length - 1);
            slideHits += hits;
            totalReplacements += hits;
          }
        } catch {
          // Shape without text frame — skip
        }
      }

      if (slideHits > 0) {
        slidesAffected.push({ slide_index: i, replacements: slideHits });
      }
    }

    await context.sync();
    return {
      search,
      replace: replace || '',
      total_replacements: totalReplacements,
      slides_affected: slidesAffected,
    };
  },

  // ─── Text Alignment ───
  async set_text_alignment(context, args) {
    const { slide_index, shape_index, alignment = 'left' } = args;
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slide_index < 0 || slide_index >= slides.items.length) {
      throw new Error(`Slide index ${slide_index} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slide_index];
    slide.shapes.load('items');
    await context.sync();

    const shapeIdx = shape_index - 1;
    if (shapeIdx < 0 || shapeIdx >= slide.shapes.items.length) {
      throw new Error(`Shape index ${shape_index} out of range (1-${slide.shapes.items.length})`);
    }

    const alignmentMap = {
      'left': 'Left',
      'center': 'Center',
      'right': 'Right',
      'justify': 'Justify',
    };

    const mappedAlignment = alignmentMap[alignment.toLowerCase()];
    if (!mappedAlignment) {
      throw new Error(`Invalid alignment '${alignment}'. Valid: left, center, right, justify`);
    }

    const shape = slide.shapes.items[shapeIdx];
    shape.textFrame.textRange.paragraphFormat.horizontalAlignment = mappedAlignment;
    await context.sync();

    return `Set alignment to ${alignment} on shape ${shape_index}, slide ${slide_index}`;
  },

  // ─── Hyperlinks ───
  async add_hyperlink(context, args) {
    const { slide_index, shape_index, url, display_text } = args;
    if (!url) throw new Error('url parameter is required');

    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();

    if (slide_index < 0 || slide_index >= slides.items.length) {
      throw new Error(`Slide index ${slide_index} out of range (0-${slides.items.length - 1})`);
    }

    const slide = slides.items[slide_index];
    slide.shapes.load('items');
    await context.sync();

    const shapeIdx = shape_index - 1;
    if (shapeIdx < 0 || shapeIdx >= slide.shapes.items.length) {
      throw new Error(`Shape index ${shape_index} out of range (1-${slide.shapes.items.length})`);
    }

    const shape = slide.shapes.items[shapeIdx];

    // Set display text if provided
    if (display_text) {
      shape.textFrame.textRange.text = display_text;
    }

    shape.textFrame.textRange.hyperlink.address = url;
    await context.sync();

    return `Hyperlink set on shape ${shape_index}, slide ${slide_index}: ${url}`;
  },

  // ─── Table Cell Formatting ───
  async format_table_cell(context, args) {
    // Table cell formatting requires PowerPointApi 1.8+ which is not widely available.
    // The table object model in Office.js is limited for PowerPoint.
    throw new Error('Table cell formatting is not yet supported via Office.js add-in (requires PowerPointApi 1.8+). Use the python-pptx MCP for table formatting or apply formatting to the whole shape via format_text.');
  },

  // ─── Presentation Creation ───
  async create_presentation(context, args) {
    throw new Error('Presentation creation is not supported via Office.js add-in (requires file system access). Use the python-pptx MCP for this operation: mcp__powerpoint__create_presentation');
  },
};
