import * as React from 'react';
import { ITinyMCEEditorProps } from '../models/IAccordionTabsModels';

// Import TinyMCE directly from node_modules
import * as tinymce from 'tinymce/tinymce';

// Import TinyMCE theme, plugins, and skin
import 'tinymce/themes/modern/theme';
import 'tinymce/plugins/advlist/plugin';
import 'tinymce/plugins/autolink/plugin';
import 'tinymce/plugins/lists/plugin';
import 'tinymce/plugins/link/plugin';
import 'tinymce/plugins/image/plugin';
import 'tinymce/plugins/charmap/plugin';
import 'tinymce/plugins/print/plugin';
import 'tinymce/plugins/preview/plugin';
import 'tinymce/plugins/anchor/plugin';
import 'tinymce/plugins/searchreplace/plugin';
import 'tinymce/plugins/visualblocks/plugin';
import 'tinymce/plugins/code/plugin';
import 'tinymce/plugins/fullscreen/plugin';
import 'tinymce/plugins/insertdatetime/plugin';
import 'tinymce/plugins/media/plugin';
import 'tinymce/plugins/table/plugin';
import 'tinymce/plugins/contextmenu/plugin';
import 'tinymce/plugins/paste/plugin';

// Import TinyMCE skin CSS
import 'tinymce/skins/lightgray/skin.min.css';
import 'tinymce/skins/lightgray/content.min.css';

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
    // TinyMCE is now imported locally and should be available
    this.setupEditor();
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
        'https://fonts.googleapis.com/css?family=Lato:300,300i,400,400i'
      ],
      // Skin CSS is now imported directly, no need to specify skin_url
      skin: false,
      // Prevent auto focus that steals focus from other fields
      auto_focus: false,
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
