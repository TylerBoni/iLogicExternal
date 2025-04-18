# iLogic External Editor Add-In

This Inventor add-in allows you to edit iLogic rules with your preferred external text editor and track them with Git.

## Features

- **External Editing**: Automatically exports iLogic rules to a folder where they can be edited with any text editor
- **Git Integration**: Rules are exported to a local 'ilogic' folder that can be tracked with Git
- **Real-time Sync**: Changes made to exported rules are automatically synced back to Inventor
- **Persistent Rules**: Rule files remain in the folder even when documents are closed
- **Selective Tracking**: Use `.ilogicignore` files to control which documents and rules are exported

## How It Works

1. Create a `.ilogicignore` file in your project folder (required)
2. The add-in creates an `ilogic` subfolder in the same location as the `.ilogicignore` file
3. Any Inventor files in that folder or subfolders will have their iLogic rules exported to the `ilogic` folder
4. You can edit these files with any text editor or track them with Git
5. Changes are automatically synced back to the Inventor document

## Using .ilogicignore Files

The add-in **requires** a `.ilogicignore` file to function. The add-in will search up the folder structure to find one.

```
# Example .ilogicignore file

# Disable transfer completely for this document
@disable-transfer

# Patterns to ignore specific rules
Test*         # Ignores all rules starting with "Test"
ExampleRule   # Ignores a specific rule named "ExampleRule"
```

If no `.ilogicignore` file is found, no rules will be tracked or exported.

### Placement Strategy

You have two options for placing your `.ilogicignore` file:

1. **Project-level**: Place a `.ilogicignore` file at the root of your project folder to track all Inventor files in the project
   - All rules will export to a single `ilogic` folder at that level
   - Good for projects with shared rules

2. **Subfolder-level**: Place a `.ilogicignore` file in a specific subfolder 
   - Only Inventor files in that subfolder will have rules exported
   - Rules will export to an `ilogic` folder in that specific subfolder
   - Good for projects with different rule sets for different components

## Git Integration

Since the rules are exported to a local `ilogic` folder in the same directory as your `.ilogicignore` file, you can easily track them with Git:

1. Initialize a Git repository in your project folder
2. Add the `ilogic/*.vb` files to your repository
3. Commit changes as you develop your rules
4. Use standard Git workflows for branching, merging, and collaboration

## Installation

1. Build the project or download the compiled DLL
2. Register the add-in with Inventor using the Add-In Manager

## Requirements

- Autodesk Inventor (tested with versions 2020+)
- iLogic module enabled in Inventor

## Development

The code is organized into several key files:

- `StandardAddInServer.cs`: The main Inventor add-in entry point
- `iLogicBridge.cs`: Core functionality for synchronizing rules
- `FileUtils.cs`: Helper utilities for file operations

## License

MIT License - See LICENSE file for details. 