# Fractal-Model-S

## Local Development Server
Run in command prompt (as Administarator) `serve.bat`.

## S3 bucket
name: fractalmodel-x-storage-us-west-2
AWS Region: us-west-2
file URL: https://fractalmodel-x-storage-us-west-2.s3-us-west-2.amazonaws.com/model.xlsm

## placement of manifest for local development
..\FractalExcelWebAddInWeb\Manifest\FractalExcelWebAddInManifestLocal.xml

## set up instruction for local development
1. Run command prompt as administrator. Open home directory of the project in it. Run `serve.bat`.
2. Use instruction from `https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins#:~:text=In%20Excel%2C%20Word%2C%20or%20PowerPoint,to%20insert%20the%20add%2Din` and manifest FractalExcelWebAddInManifestLocal.xml (placement: `..\FractalExcelWebAddInWeb\Manifest\FractalExcelWebAddInManifestLocal.xml`) to add add-in to Excel.
3. Open Excel.
4. Run Add-in from Excel control panel.

## troubleshooting
If 'loading' label is dislaying too long in add-in frame:
  1. Close add-in.
  2. Stop script `serve.bat`(`Ctrl + Break` in command prompt and then accept `Y`).
  3. Run script `serve.bat`.
  4. Run Add-in from Excel control panel.

## Deployment ##
Prerequisites:
 - node/npm
 - python3
To deploy addin to `addin.fractalmodel.com` run `deploy/deployAddin.sh` script.
P.S. Since this is a `bash` (`sh`) script it cannot be run on Windows natively. To run it you can use `WSL` or `Git Bash` terminal.