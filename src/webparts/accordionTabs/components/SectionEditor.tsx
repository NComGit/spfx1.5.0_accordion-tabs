import * as React from 'react';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { ISectionEditorProps } from '../models/IAccordionTabsModels';
import { ISection } from '../models/IAccordionTabsModels';
import { TinyMCEEditor } from './TinyMCEEditor';
import styles from './SectionEditor.module.scss';

interface ISectionEditorState {
  title: string;
  content: string;
  isLoading: boolean;
  hasChanges: boolean;
}

/**
 * Section Editor component for editing individual accordion/tab sections
 * Uses TinyMCE 4.x for rich text editing
 */
export class SectionEditor extends React.Component<ISectionEditorProps, ISectionEditorState> {
  
  constructor(props: ISectionEditorProps) {
    super(props);

    this.state = {
      title: props.section ? props.section.title : '',
      content: props.section ? props.section.content : '',
      isLoading: false,
      hasChanges: false
    };

    // Bind methods
    this.onTitleChange = this.onTitleChange.bind(this);
    this.onContentChange = this.onContentChange.bind(this);
    this.onSave = this.onSave.bind(this);
    this.onCancel = this.onCancel.bind(this);
  }

  public componentDidUpdate(prevProps: ISectionEditorProps): void {
    // Update state when section prop changes
    if (prevProps.section !== this.props.section) {
      this.setState((prevState) => ({
        ...prevState,
        title: this.props.section ? this.props.section.title : '',
        content: this.props.section ? this.props.section.content : '',
        hasChanges: false
      }));
    }
  }

  private onTitleChange(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void {
    this.setState((prevState) => ({
      ...prevState,
      title: newValue || '',
      hasChanges: true
    }));
  }

  private onContentChange(content: string): void {
    this.setState((prevState) => ({
      ...prevState,
      content: content,
      hasChanges: true
    }));
  }

  private onSave(): void {
    const { title, content } = this.state;
    const { section, onSave } = this.props;

    if (!title.trim()) {
      // Could add error handling here
      return;
    }

    this.setState((prevState) => ({
      ...prevState,
      isLoading: true
    }));

    const updatedSection: ISection = {
      id: section ? section.id : this.generateId(),
      title: title.trim(),
      content: content,
      order: section ? section.order : 0
    };

    // Simulate async operation
    setTimeout(() => {
      this.setState((prevState) => ({
        ...prevState,
        isLoading: false,
        hasChanges: false
      }));
      onSave(updatedSection);
    }, 500);
  }

  private onCancel(): void {
    const { onCancel } = this.props;
    
    // Reset state to original values
    this.setState((prevState) => ({
      ...prevState,
      title: this.props.section ? this.props.section.title : '',
      content: this.props.section ? this.props.section.content : '',
      hasChanges: false
    }));

    onCancel();
  }

  private generateId(): string {
    return 'section-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
  }

  public render(): React.ReactElement<ISectionEditorProps> {
    const { isVisible, section } = this.props;
    const { title, content, isLoading, hasChanges } = this.state;
    
    const modalTitle = section ? 'Edit Section' : 'Add New Section';

    return (
      <Modal
        isOpen={isVisible}
        onDismiss={this.onCancel}
        isBlocking={hasChanges}
        containerClassName={styles.modalContainer}
      >
        <div className={styles.modalHeader}>
          <h2 className={styles.modalTitle}>{modalTitle}</h2>
        </div>
        
        <div className={styles.modalBody}>
          <div className={styles.fieldGroup}>
            <TextField
              label="Section Title"
              value={title}
              onChange={this.onTitleChange}
              required
              disabled={isLoading}
              placeholder="Enter section title..."
            />
          </div>

          <div className={styles.fieldGroup}>
            <label className={styles.fieldLabel}>Section Content</label>
            <div className={styles.editorContainer}>
              <TinyMCEEditor
                value={content}
                onEditorChange={this.onContentChange}
                height={300}
              />
            </div>
          </div>
        </div>

        <div className={styles.modalFooter}>
          <PrimaryButton
            text={section ? 'Update Section' : 'Add Section'}
            onClick={this.onSave}
            disabled={isLoading || !title.trim()}
          />
          <DefaultButton
            text="Cancel"
            onClick={this.onCancel}
            disabled={isLoading}
            style={{ marginLeft: '10px' }}
          />
        </div>

        {isLoading && (
          <div className={styles.loadingOverlay}>
            <div className={styles.spinner}>Saving...</div>
          </div>
        )}
      </Modal>
    );
  }
}
