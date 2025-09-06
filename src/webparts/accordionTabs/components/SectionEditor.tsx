import * as React from 'react';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { ISectionEditorProps, ISectionEditorState, ISection } from '../models/IAccordionTabsModels';
import { TinyMCEEditor } from './TinyMCEEditor';
import styles from './SectionEditor.module.scss';

/**
 * Section Editor component for editing individual accordion/tab sections
 * Uses TinyMCE 4.x for rich text editing
 */
export class SectionEditor extends React.Component<ISectionEditorProps, ISectionEditorState> {
  
  constructor(props: ISectionEditorProps) {
    super(props);

    const initialTitle = props.section ? props.section.title : '';
    const initialContent = props.section ? props.section.content : '';
    

    this.state = {
      title: initialTitle,
      content: initialContent,
      isLoading: false,
      hasChanges: false
    };

    // Bind methods
    this.onTitleChange = this.onTitleChange.bind(this);
    this.onTitleChangeSimple = this.onTitleChangeSimple.bind(this);
    this.onContentChange = this.onContentChange.bind(this);
    this.onSave = this.onSave.bind(this);
    this.onCancel = this.onCancel.bind(this);
    this.onTitleFocus = this.onTitleFocus.bind(this);
    this.onTitleBlur = this.onTitleBlur.bind(this);
  }

  private onTitleFocus(): void {
    // Focus handler for title field
  }

  private onTitleBlur(): void {
    // Blur handler for title field
  }

  public componentDidUpdate(prevProps: ISectionEditorProps): void {
    // Only update state when we're actually switching between different sections
    // or when opening the modal for the first time
    const prevSectionId = prevProps.section ? prevProps.section.id : 'new';
    const currentSectionId = this.props.section ? this.props.section.id : 'new';
    const modalJustOpened = !prevProps.isVisible && this.props.isVisible;
    const sectionChanged = prevSectionId !== currentSectionId;
    
    if (modalJustOpened || sectionChanged) {
      const newTitle = this.props.section ? this.props.section.title : '';
      const newContent = this.props.section ? this.props.section.content : '';
      
      this.setState((prevState) => ({
        ...prevState,
        title: newTitle,
        content: newContent,
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

  private onTitleChangeSimple(newValue: string): void {
    this.setState((prevState) => ({
      ...prevState,
      title: newValue,
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
            <label className={styles.fieldLabel}>Section Title *</label>
            <input
              type="text"
              value={title}
              onChange={(e) => this.onTitleChangeSimple(e.target.value)}
              onFocus={this.onTitleFocus}
              onBlur={this.onTitleBlur}
              disabled={isLoading}
              placeholder="Enter section title..."
              style={{
                width: '100%',
                padding: '8px 12px',
                border: '1px solid #a6a6a6',
                borderRadius: '2px',
                fontSize: '14px',
                outline: 'none'
              }}
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
