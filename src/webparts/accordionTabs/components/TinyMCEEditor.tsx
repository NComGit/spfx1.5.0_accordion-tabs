import * as React from 'react';
import { ITinyMCEEditorProps } from '../models/IAccordionTabsModels';

// Import TinyMCE from npm package instead of CDN
import 'tinymce/tinymce.min.js';

// Import required plugins
import 'tinymce/themes/modern/theme.js';
import 'tinymce/plugins/lists/plugin.js';
import 'tinymce/plugins/textcolor/plugin.js';
import 'tinymce/plugins/colorpicker/plugin.js';
import 'tinymce/plugins/link/plugin.js';
import 'tinymce/plugins/image/plugin.js';
import 'tinymce/plugins/paste/plugin.js';

// Import skin CSS
import 'tinymce/skins/lightgray/skin.min.css';
import 'tinymce/skins/lightgray/content.min.css';

declare var tinymce: any;

/**
 * TinyMCE 4.x Editor component for React 15.6.2
 * Simplified configuration for better compatibility and functionality
 */
export class TinyMCEEditor extends React.Component<ITinyMCEEditorProps, {}> {
  private editorId: string;
  private editor: any;

  constructor(props: ITinyMCEEditorProps) {
    super(props);
    this.editorId = 'tinymce-editor-' + Math.random().toString(36).substr(2, 9);
  }

  public componentDidMount(): void {
    this.initializeTinyMCE();
  }

  public componentWillUnmount(): void {
    if (this.editor) {
      this.editor.remove();
    }
  }

  public componentDidUpdate(prevProps: ITinyMCEEditorProps): void {
    // Update editor content if value changed externally
    if (prevProps.value !== this.props.value && this.editor) {
      const currentContent = this.editor.getContent();
      
      if (currentContent !== this.props.value) {
        this.editor.setContent(this.props.value || '');
      }
    }
  }

  private showFormatDropdown(editor: any, formats: string[], formatNames: string[]): void {
    // Remove any existing dropdown
    const existingDropdown = document.getElementById('tinymce-format-dropdown');
    if (existingDropdown) {
      document.body.removeChild(existingDropdown);
    }

    // Find the button position
    const editorContainer = editor.getContainer();
    const formatButton = editorContainer.querySelector('[aria-label="Select format"]') || editorContainer.querySelector('button[title*="format" i]');
    
    if (!formatButton) {
      console.log('Format button not found, falling back to editor position');
    }

    // Create dropdown
    const dropdown = document.createElement('div');
    dropdown.id = 'tinymce-format-dropdown';
    dropdown.style.cssText = 'position: absolute; background: white; border: 1px solid #ccc; border-radius: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); min-width: 180px; z-index: 2147483647; max-height: 200px; overflow-y: auto;';

    // Position dropdown near the button or editor
    if (formatButton) {
      const buttonRect = formatButton.getBoundingClientRect();
      dropdown.style.top = (buttonRect.bottom + window.scrollY + 2) + 'px';
      dropdown.style.left = (buttonRect.left + window.scrollX) + 'px';
    } else {
      const editorRect = editorContainer.getBoundingClientRect();
      dropdown.style.top = (editorRect.top + window.scrollY - 200) + 'px';
      dropdown.style.left = (editorRect.left + window.scrollX) + 'px';
    }

    // Add options
    for (let i = 0; i < formats.length; i++) {
      const option = document.createElement('div');
      option.textContent = formatNames[i];
      option.style.cssText = 'padding: 8px 12px; cursor: pointer; border-bottom: 1px solid #f0f0f0;';
      option.onmouseover = () => option.style.backgroundColor = '#f0f0f0';
      option.onmouseout = () => option.style.backgroundColor = 'white';
      option.onclick = () => {
        // Check if there's a selection
        const selection = editor.selection.getContent();
        if (selection && selection.trim()) {
          // For selected text, wrap in the appropriate tag
          const tagName = formats[i];
          if (tagName === 'p') {
            // For paragraph, just remove any existing formatting
            editor.execCommand('RemoveFormat', false);
          } else {
            // For headings, wrap the selection
            editor.execCommand('mceInsertContent', false, `<${tagName}>${selection}</${tagName}>`);
          }
        } else {
          // For no selection, format the current block
          editor.execCommand('FormatBlock', false, formats[i]);
        }
        document.body.removeChild(dropdown);
        console.log('Applied format:', formatNames[i]);
      };
      dropdown.appendChild(option);
    }

    // Close dropdown when clicking outside
    const closeDropdown = (e: Event) => {
      if (!dropdown.contains(e.target as Node) && document.body.contains(dropdown)) {
        document.body.removeChild(dropdown);
        document.removeEventListener('click', closeDropdown);
      }
    };

    document.body.appendChild(dropdown);
    setTimeout(() => document.addEventListener('click', closeDropdown), 10);
  }

  private showSizeDropdown(editor: any, sizes: string[], sizeNames: string[]): void {
    // Remove any existing dropdown
    const existingDropdown = document.getElementById('tinymce-size-dropdown');
    if (existingDropdown) {
      document.body.removeChild(existingDropdown);
    }

    // Find the button position
    const editorContainer = editor.getContainer();
    const sizeButton = editorContainer.querySelector('[aria-label="Select font size"]') || editorContainer.querySelector('button[title*="size" i]');
    
    if (!sizeButton) {
      console.log('Size button not found, falling back to editor position');
    }

    // Create dropdown
    const dropdown = document.createElement('div');
    dropdown.id = 'tinymce-size-dropdown';
    dropdown.style.cssText = 'position: absolute; background: white; border: 1px solid #ccc; border-radius: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); min-width: 120px; z-index: 2147483647; max-height: 150px; overflow-y: auto;';

    // Position dropdown near the button or editor
    if (sizeButton) {
      const buttonRect = sizeButton.getBoundingClientRect();
      dropdown.style.top = (buttonRect.bottom + window.scrollY + 2) + 'px';
      dropdown.style.left = (buttonRect.left + window.scrollX) + 'px';
    } else {
      const editorRect = editorContainer.getBoundingClientRect();
      dropdown.style.top = (editorRect.top + window.scrollY - 150) + 'px';
      dropdown.style.left = (editorRect.left + window.scrollX + 50) + 'px';
    }

    // Add options
    for (let i = 0; i < sizes.length; i++) {
      const option = document.createElement('div');
      option.textContent = sizeNames[i];
      option.style.cssText = 'padding: 6px 12px; cursor: pointer; border-bottom: 1px solid #f0f0f0;';
      option.onmouseover = () => option.style.backgroundColor = '#f0f0f0';
      option.onmouseout = () => option.style.backgroundColor = 'white';
      option.onclick = () => {
        editor.execCommand('FontSize', false, sizes[i]);
        document.body.removeChild(dropdown);
        console.log('Applied font size:', sizeNames[i]);
      };
      dropdown.appendChild(option);
    }

    // Close dropdown when clicking outside
    const closeDropdown = (e: Event) => {
      if (!dropdown.contains(e.target as Node) && document.body.contains(dropdown)) {
        document.body.removeChild(dropdown);
        document.removeEventListener('click', closeDropdown);
      }
    };

    document.body.appendChild(dropdown);
    setTimeout(() => document.addEventListener('click', closeDropdown), 10);
  }

  private showTextColorDropdown(editor: any, textColors: string[], textColorNames: string[]): void {
    // Remove any existing dropdown
    const existingDropdown = document.getElementById('tinymce-textcolor-dropdown');
    if (existingDropdown) {
      document.body.removeChild(existingDropdown);
    }

    // Find the button position
    const editorContainer = editor.getContainer();
    const colorButton = editorContainer.querySelector('[aria-label="Select text color"]') || editorContainer.querySelector('button[title*="text color" i]');
    
    if (!colorButton) {
      console.log('Text color button not found, falling back to editor position');
    }

    // Create dropdown
    const dropdown = document.createElement('div');
    dropdown.id = 'tinymce-textcolor-dropdown';
    dropdown.style.cssText = 'position: absolute; background: white; border: 1px solid #ccc; border-radius: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); min-width: 160px; z-index: 2147483647; max-height: 180px; overflow-y: auto;';

    // Position dropdown near the button or editor
    if (colorButton) {
      const buttonRect = colorButton.getBoundingClientRect();
      dropdown.style.top = (buttonRect.bottom + window.scrollY + 2) + 'px';
      dropdown.style.left = (buttonRect.left + window.scrollX) + 'px';
    } else {
      const editorRect = editorContainer.getBoundingClientRect();
      dropdown.style.top = (editorRect.top + window.scrollY - 180) + 'px';
      dropdown.style.left = (editorRect.left + window.scrollX + 100) + 'px';
    }

    // Add options
    for (let i = 0; i < textColors.length; i++) {
      const option = document.createElement('div');
      option.innerHTML = `<span style="background-color: ${textColors[i]}; width: 16px; height: 16px; display: inline-block; margin-right: 8px; border: 1px solid #ccc; border-radius: 2px;"></span>${textColorNames[i]}`;
      option.style.cssText = 'padding: 6px 12px; cursor: pointer; border-bottom: 1px solid #f0f0f0; display: flex; align-items: center;';
      option.onmouseover = () => option.style.backgroundColor = '#f0f0f0';
      option.onmouseout = () => option.style.backgroundColor = 'white';
      option.onclick = () => {
        editor.execCommand('ForeColor', false, textColors[i]);
        document.body.removeChild(dropdown);
        console.log('Applied text color:', textColorNames[i]);
      };
      dropdown.appendChild(option);
    }

    // Close dropdown when clicking outside
    const closeDropdown = (e: Event) => {
      if (!dropdown.contains(e.target as Node) && document.body.contains(dropdown)) {
        document.body.removeChild(dropdown);
        document.removeEventListener('click', closeDropdown);
      }
    };

    document.body.appendChild(dropdown);
    setTimeout(() => document.addEventListener('click', closeDropdown), 10);
  }

  private showBgColorDropdown(editor: any, bgColors: string[], bgColorNames: string[]): void {
    // Remove any existing dropdown
    const existingDropdown = document.getElementById('tinymce-bgcolor-dropdown');
    if (existingDropdown) {
      document.body.removeChild(existingDropdown);
    }

    // Find the button position
    const editorContainer = editor.getContainer();
    const bgButton = editorContainer.querySelector('[aria-label="Select background color"]') || editorContainer.querySelector('button[title*="background" i]');
    
    if (!bgButton) {
      console.log('Background color button not found, falling back to editor position');
    }

    // Create dropdown
    const dropdown = document.createElement('div');
    dropdown.id = 'tinymce-bgcolor-dropdown';
    dropdown.style.cssText = 'position: absolute; background: white; border: 1px solid #ccc; border-radius: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); min-width: 180px; z-index: 2147483647; max-height: 180px; overflow-y: auto;';

    // Position dropdown near the button or editor
    if (bgButton) {
      const buttonRect = bgButton.getBoundingClientRect();
      dropdown.style.top = (buttonRect.bottom + window.scrollY + 2) + 'px';
      dropdown.style.left = (buttonRect.left + window.scrollX) + 'px';
    } else {
      const editorRect = editorContainer.getBoundingClientRect();
      dropdown.style.top = (editorRect.top + window.scrollY - 180) + 'px';
      dropdown.style.left = (editorRect.left + window.scrollX + 150) + 'px';
    }

    // Add options
    for (let i = 0; i < bgColors.length; i++) {
      const option = document.createElement('div');
      const colorDisplay = bgColors[i] ? `<span style="background-color: ${bgColors[i]}; width: 16px; height: 16px; display: inline-block; margin-right: 8px; border: 1px solid #ccc; border-radius: 2px;"></span>` : '<span style="width: 16px; height: 16px; display: inline-block; margin-right: 8px; border: 1px solid #ccc; background: white; position: relative; border-radius: 2px;"><span style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); font-size: 10px;">×</span></span>';
      option.innerHTML = `${colorDisplay}${bgColorNames[i]}`;
      option.style.cssText = 'padding: 6px 12px; cursor: pointer; border-bottom: 1px solid #f0f0f0; display: flex; align-items: center;';
      option.onmouseover = () => option.style.backgroundColor = '#f0f0f0';
      option.onmouseout = () => option.style.backgroundColor = 'white';
      option.onclick = () => {
        if (bgColors[i] === '') {
          // For "None", remove background color by setting it to transparent
          editor.execCommand('BackColor', false, 'transparent');
          // Also try removing the style attribute for background color
          const selection = editor.selection.getNode();
          if (selection && selection.style) {
            selection.style.backgroundColor = '';
          }
        } else {
          editor.execCommand('BackColor', false, bgColors[i]);
        }
        document.body.removeChild(dropdown);
        console.log('Applied background color:', bgColorNames[i]);
      };
      dropdown.appendChild(option);
    }

    // Close dropdown when clicking outside
    const closeDropdown = (e: Event) => {
      if (!dropdown.contains(e.target as Node) && document.body.contains(dropdown)) {
        document.body.removeChild(dropdown);
        document.removeEventListener('click', closeDropdown);
      }
    };

    document.body.appendChild(dropdown);
    setTimeout(() => document.addEventListener('click', closeDropdown), 10);
  }


  private initializeTinyMCE(): void {
    // TinyMCE is now loaded via npm imports, so directly setup editor
    console.log('Initializing TinyMCE from npm package');
    this.setupEditor();
  }

  private setupEditor(): void {
    const { height = 300, onEditorChange, value } = this.props;
    
    // Add debugging
    console.log('Setting up TinyMCE editor with ID:', this.editorId);
    
    // State variables for cycling buttons
    let currentFormat = 0;
    const formats = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
    const formatNames = ['Paragraph', 'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Heading 5', 'Heading 6'];
    
    let currentSize = 3;
    const sizes = ['1', '2', '3', '4', '5', '6', '7'];
    const sizeNames = ['8px', '10px', '12px', '14px', '18px', '24px', '36px'];
    
    let currentTextColor = 0;
    const textColors = ['#000000', '#ff0000', '#008000', '#0000ff', '#ffa500', '#800080'];
    const textColorNames = ['Black', 'Red', 'Green', 'Blue', 'Orange', 'Purple'];
    
    let currentBgColor = 0;
    const bgColors = ['', '#ffff00', '#add8e6', '#90ee90', '#ffb6c1', '#d3d3d3'];
    const bgColorNames = ['None', 'Yellow', 'Light Blue', 'Light Green', 'Light Pink', 'Light Gray'];

    tinymce.init({
      selector: `#${this.editorId}`,
      height: height,
      menubar: false,
      statusbar: false,
      
      // Prevent TinyMCE from loading external resources
      skin: false,
      content_css: false,
      
      // Fix plugin loading - these are the exact plugins needed for each button (removed 'code' plugin)
      plugins: 'lists textcolor colorpicker link image paste',
      
      // Updated toolbar using custom buttons that actually work
      toolbar: 'undo redo | customformat customfontsize | bold italic underline | customtextcolor custombgcolor | alignleft aligncenter alignright | bullist numlist | customimage',
      
      // Remove formatselect and fontsizeselect temporarily to focus on color/image/code
      // Font size options that should work
      fontsize_formats: '8pt 10pt 12pt 14pt 16pt 18pt 24pt 36pt',
      
      // Format dropdown options
      block_formats: 'Paragraph=p; Heading 1=h1; Heading 2=h2; Heading 3=h3; Heading 4=h4; Heading 5=h5; Heading 6=h6',
      
      // Ensure textcolor plugin works properly
      textcolor_cols: 5,
      textcolor_rows: 5,
      
      // Basic settings for SharePoint compatibility
      auto_focus: false,
      browser_spellcheck: true,
      forced_root_block: false,
      force_br_newlines: true,
      force_p_newlines: false,
      convert_urls: false,
      remove_script_host: false,
      
      // File picker for images
      file_picker_types: 'image',
      file_picker_callback: (cb: any, pickerVal: any, meta: any) => {
        console.log('Image picker called');
        const input = document.createElement('input');
        input.setAttribute('type', 'file');
        input.setAttribute('accept', 'image/*');
        
        input.onchange = (event: Event) => {
          const target = event.target as HTMLInputElement;
          const file = target.files && target.files[0];
          if (file) {
            const reader = new FileReader();
            reader.onload = () => {
              cb(reader.result, { alt: file.name });
            };
            reader.readAsDataURL(file);
          }
        };
        
        input.click();
      },
      
      // Setup editor with proper event handling and debugging
      setup: (editor: any) => {
        this.editor = editor;
        
        // Custom format button with dropdown functionality
        editor.addButton('customformat', {
          text: 'Format',
          tooltip: 'Select format',
          onclick: () => {
            console.log('Custom format button clicked');
            this.showFormatDropdown(editor, formats, formatNames);
          }
        });

        // Custom font size button with dropdown functionality
        editor.addButton('customfontsize', {
          text: 'Size',
          tooltip: 'Select font size',
          onclick: () => {
            console.log('Custom font size button clicked');
            this.showSizeDropdown(editor, sizes, sizeNames);
          }
        });

        editor.addButton('customtextcolor', {
          icon: 'forecolor',
          tooltip: 'Select text color',
          onclick: () => {
            console.log('Custom text color button clicked');
            this.showTextColorDropdown(editor, textColors, textColorNames);
          }
        });

        editor.addButton('custombgcolor', {
          icon: 'backcolor',
          tooltip: 'Select background color',
          onclick: () => {
            console.log('Custom bg color button clicked');
            this.showBgColorDropdown(editor, bgColors, bgColorNames);
          }
        });

        editor.addButton('customimage', {
          icon: 'image',
          tooltip: 'Insert Image',
          onclick: () => {
            console.log('Custom image button clicked');
            const input = document.createElement('input');
            input.setAttribute('type', 'file');
            input.setAttribute('accept', 'image/*');
            
            input.onchange = (event: Event) => {
              const target = event.target as HTMLInputElement;
              const file = target.files && target.files[0];
              if (file) {
                const reader = new FileReader();
                reader.onload = () => {
                  editor.execCommand('mceInsertContent', false, `<img src="${reader.result}" alt="${file.name}" style="max-width: 100%;" />`);
                };
                reader.readAsDataURL(file);
              }
            };
            
            input.click();
          }
        });


        // Initialize content when editor is ready
        editor.on('init', () => {
          console.log('TinyMCE editor initialized successfully');
          console.log('Available buttons:', editor.buttons);
          editor.setContent(value || '');
        });

        // Add debugging for button clicks
        editor.on('click', (e: any) => {
          console.log('Editor clicked:', e.target);
        });

        // Debug toolbar button clicks
        editor.on('BeforeExecCommand', (e: any) => {
          console.log('Command executed:', e.command, 'Value:', e.value);
        });

        // Handle content changes
        editor.on('change', () => {
          const content = editor.getContent();
          if (onEditorChange) {
            onEditorChange(content);
          }
        });

        // Handle keyup for real-time updates
        editor.on('keyup', () => {
          const content = editor.getContent();
          if (onEditorChange) {
            onEditorChange(content);
          }
        });

        // Handle paste events
        editor.on('paste', () => {
          setTimeout(() => {
            const content = editor.getContent();
            if (onEditorChange) {
              onEditorChange(content);
            }
          }, 100);
        });

        // Handle blur to ensure final save
        editor.on('blur', () => {
          const content = editor.getContent();
          if (onEditorChange) {
            onEditorChange(content);
          }
        });
      }
    });
  }

  public render(): React.ReactElement<ITinyMCEEditorProps> {
    return (
      <div>
        <textarea id={this.editorId} style={{ width: '100%', minHeight: '200px' }} />
      </div>
    );
  }
}
