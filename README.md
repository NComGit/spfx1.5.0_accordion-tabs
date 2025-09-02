# Accordion/Tabs Web Part for SharePoint Framework 1.5.0

A dynamic SharePoint Framework (SPFx) web part that displays content in either accordion or tabs format, featuring rich text editing capabilities with TinyMCE 4.x integration.

## Features

- **Dual View Modes**: Switch between accordion and tabs display formats
- **Rich Text Editing**: TinyMCE 4.x integration for content editing with full formatting support
- **Dynamic Content Management**: Add, edit, delete, and reorder sections without limitations
- **Responsive Design**: Optimized for desktop, tablet, and mobile devices
- **SharePoint Integration**: Full integration with SharePoint property pane
- **TypeScript Support**: Built with TypeScript 2.4.2 for type safety
- **Office UI Fabric**: Consistent SharePoint look and feel

## Technical Specifications

- **SPFx Version**: 1.5.0
- **Target Platform**: SharePoint Server (on-premises)
- **React Version**: 15.6.2 (Class Components)
- **TypeScript Version**: ~2.4.2
- **TinyMCE Version**: 4.9.11 (latest 4.x compatible with React 15.6.2)
- **Office UI Fabric React**: 5.131.0

## Prerequisites

- Node.js 8.11.0 or higher
- SharePoint Framework development environment
- SharePoint Server (on-premises) environment

## Installation

### 1. Install Dependencies

```bash
npm install
```

### 2. Build the Solution

```bash
# Development build
gulp build

# Production build
gulp bundle --ship
gulp package-solution --ship
```

### 3. Deploy to SharePoint

1. Upload the `.sppkg` file from `sharepoint/solution/` to your SharePoint App Catalog
2. Install the app in your SharePoint site
3. Add the web part to a page

## Project Structure

```
src/
├── webparts/
│   └── accordionTabs/
│       ├── AccordionTabsWebPart.ts           # Main web part class
│       ├── AccordionTabsWebPart.module.scss  # Web part styles
│       ├── models/
│       │   └── IAccordionTabsModels.ts       # TypeScript interfaces
│       ├── components/
│       │   ├── AccordionTabsComponent.tsx    # Main React component
│       │   ├── AccordionView.tsx             # Accordion view component
│       │   ├── TabsView.tsx                  # Tabs view component
│       │   ├── SectionEditor.tsx             # Section editor modal
│       │   ├── TinyMCEEditor.tsx             # TinyMCE integration
│       │   └── *.module.scss                 # Component styles
│       └── loc/
│           ├── mystrings.d.ts                # Localization interface
│           └── en-us.js                      # English strings
```

## Configuration

### Property Pane Settings

1. **Display Settings**
   - View Type: Choose between Accordion or Tabs layout

2. **Content Sections**
   - Add New Section: Create additional content sections
   - Section Management: Edit titles, reorder, and delete sections

### Content Management

**Edit Mode Features:**
- Click "Add Section" to create new content areas
- Use edit buttons to modify section titles and content
- Drag and reorder sections using up/down arrows
- Delete unwanted sections

**Rich Text Editor:**
- Full TinyMCE 4.x integration
- Support for bold, italic, underline, lists, links
- Image insertion and basic formatting
- HTML content storage

## Usage Examples

### Basic Implementation

The web part automatically renders based on the configured view type:

**Accordion View:**
- Collapsible sections with expand/collapse functionality
- First section expanded by default
- Click headers to toggle content visibility

**Tabs View:**
- Horizontal tab navigation
- Active tab highlighting
- Click tabs to switch content

### Content Structure

Each section contains:
- **Title**: Plain text identifier
- **Content**: Rich HTML content from TinyMCE
- **Order**: Position in the display sequence

## Customization

### Styling

Modify SCSS files to customize appearance:
- `AccordionView.module.scss`: Accordion-specific styles
- `TabsView.module.scss`: Tabs-specific styles
- `SectionEditor.module.scss`: Editor modal styles

### Localization

Add new language support:
1. Create new language file in `loc/` (e.g., `fr-fr.js`)
2. Update `mystrings.d.ts` if adding new strings
3. Configure in SharePoint for multi-language sites

### TinyMCE Configuration

Customize the editor in `TinyMCEEditor.tsx`:
- Modify toolbar options
- Add/remove plugins
- Adjust editor height
- Configure content filtering

## Browser Support

- Microsoft Edge
- Chrome (latest)
- Firefox (latest)
- Safari (latest)

## Development

### Debug Mode

```bash
gulp serve
```

### Building for Production

```bash
gulp clean
gulp bundle --ship
gulp package-solution --ship
```

## Troubleshooting

### Common Issues

1. **TinyMCE Not Loading**
   - Verify CDN access to TinyMCE scripts
   - Check browser console for script loading errors

2. **Property Pane Issues**
   - Ensure all localization strings are defined
   - Verify property pane field names match interface

3. **Styling Problems**
   - Check SCSS compilation
   - Verify Office UI Fabric dependencies

4. **TypeScript Errors**
   - Ensure compatibility with TypeScript 2.4.2
   - Use manual loops instead of array methods for compatibility

### Error Handling

The web part includes comprehensive error handling:
- Loading states during content operations
- Validation for required fields
- Graceful degradation for missing content

## Performance Considerations

- TinyMCE loads asynchronously to prevent blocking
- Efficient React component rendering
- Minimal DOM manipulation
- Optimized SCSS compilation

## Security

- Content sanitization through TinyMCE
- SharePoint permission integration
- No external data storage
- HTTPS enforcement for CDN resources

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes following TypeScript/React best practices
4. Test thoroughly in SharePoint environment
5. Submit pull request with detailed description


