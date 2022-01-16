# Excel AddIn Angular

Excel AddIn using Angular

## Notes

- Notes
  - Creating from scratch

    ```s
      # install Yeoman generator for Office Add-ins
      npm install -g yo generator-office

      # (optional) opt out Office Add-in CLI tools anonymized usage data collection
      npx office-addin-usage-data off

      # scaffold the project
      yo office --projectType angular --name "Notion API AddIn" --host excel --ts true

      # change the directory to the project created and start with npm start
      npm start

      # Accept the prompt to run the add-ins over HTTPS
      # Debugging is being started...
      # App type: desktop
      # The developer certificates have been generated in /Users/{user-name}/.office-addin-dev-certs
      # Installing CA certificate "Developer CA for Microsoft Office Add-ins"...
      # Password:
    ```
  
  - Development
    - Windows
      - Debug
        - Run: `npm start`
          - A `Node.js JavaScript Runtime` window should pop up running `webpack`
          - An `Excel` window should open with the add-in in the top right `Commands Group` window
        - Click on the add-in
        - At the `WebView Stop On Load` prompt it will inform you to attach VS Code to the webview instance to debug, start the `Excel Desktop (Chrome)` launch config and press `OK`
        - Stop: `npm run stop`

  - Issues
    - Would not start with `npm start`
      - Error

        ```log
          > office-addin-taskpane-angular@0.0.1 start /Volumes/AppData/Programming/JavaScript/MSOffice/Excel/Notion API AddIn
          > office-addin-debugging start manifest.xml

          Debugging is being started...
          App type: desktop
          The dev server is already running on port 3000.
          Sideloading the Office Add-in...
          Error: Unable to start debugging.
          Error: Unable to sideload the Office Add-in. 
          Error: Unable to register the Office Add-in.
          Error: EXDEV: cross-device link not permitted, link 'manifest.xml' -> '/Users/delemar/Library/Containers/com.microsoft.Excel/Data/Documents/wef/932f2091-e778-4b46-9675-4d51991a6463.manifest.xml'
        ```

### References

- References
  - [Yeoman generator for Office Add-ins - YO OFFICE!](https://github.com/OfficeDev/generator-office)
