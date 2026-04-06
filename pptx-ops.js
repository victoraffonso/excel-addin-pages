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

  // ─── Presentation Creation ───
  async create_presentation(context, args) {
    throw new Error('Presentation creation is not supported via Office.js add-in (requires file system access). Use the python-pptx MCP for this operation: mcp__powerpoint__create_presentation');
  },
};
