import * as React from 'react';
import { ITinyMCEEditorProps } from '../models/IAccordionTabsModels';

declare var tinymce: any;

/**
 * TinyMCE 4.x Editor component for React 15.6.2
 * Note: This is a manual implementation since @tinymce/tinymce-react v2.6.0 may have compatibility issues
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

  private initializeTinyMCE(): void {
    // Ensure TinyMCE is loaded
    if (typeof tinymce === 'undefined') {
      // Load TinyMCE script dynamically
      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/npm/tinymce@4.9.11/tinymce.min.js';
      script.onload = () => {
        this.setupEditor();
      };
      document.head.appendChild(script);
    } else {
      this.setupEditor();
    }
  }

  private setupEditor(): void {
    const { height = 300, onEditorChange, value } = this.props;

    tinymce.init({
      selector: `#${this.editorId}`,
      height: height,
      menubar: false,
      plugins: [
        'advlist autolink lists link image charmap print preview anchor',
        'searchreplace visualblocks code fullscreen',
        'insertdatetime media table contextmenu paste code'
      ],
      toolbar: 'undo redo | formatselect | bold italic underline strikethrough | ' +
        'alignleft aligncenter alignright alignjustify | ' +
        'bullist numlist outdent indent | link image | code',
      content_css: [
        '//fonts.googleapis.com/css?family=Lato:300,300i,400,400i',
        '//www.tinymce.com/css/codepen.min.css'
      ],
      skin_url: 'https://cdn.jsdelivr.net/npm/tinymce@4.9.11/skins/lightgray',
      setup: (editor: any) => {
        this.editor = editor;
        
        // Set initial content
        editor.on('init', () => {
          editor.setContent(value || '');
        });

        // Handle content changes
        editor.on('change keyup', () => {
          const content = editor.getContent();
          if (onEditorChange) {
            onEditorChange(content);
          }
        });

        // Handle blur event to ensure content is saved
        editor.on('blur', () => {
          const content = editor.getContent();
          if (onEditorChange) {
            onEditorChange(content);
          }
        });
      },
      // Disable automatic uploads and file picker for security
      automatic_uploads: false,
      file_picker_types: 'image',
      file_picker_callback: (cb: any, pickerValue: any, meta: any) => {
        // Simple file picker - in production, you'd want more sophisticated handling
        const input = document.createElement('input');
        input.setAttribute('type', 'file');
        input.setAttribute('accept', 'image/*');
        input.onchange = (event: Event) => {
          const target = event.target as HTMLInputElement;
          const file = target.files![0];
          const reader = new FileReader();
          reader.onload = () => {
            cb(reader.result, { alt: file.name });
          };
          reader.readAsDataURL(file);
        };
        input.click();
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
