# Workspace

This folder contains application-managed state. The application creates missing subfolders automatically.

```text
workspace/
|-- templates/     Saved PowerPoint templates and template images
|-- mappings/      Platform configurations and PowerPoint mappings
|-- logs/          Rotating application logs
`-- settings.json  Created after settings are saved
```

The user-facing `input/` and `output/` folders are located directly in the project root folder for easier navigation.

Back up `templates/`, `mappings/`, and `settings.json` when moving the application to another machine. Reports in the root `output/` folder can be regenerated from the original exports.
