# medibrane-aggregations

### commands

once - `gulp trust-dev-cert`

`gulp serve`

`gulp build`
`gulp bundle --ship`
`gulp package-solution --ship`

https://colorlib.com/wp/free-html5-admin-dashboard-templates/




https://{tenant}/sites/{subsite}/_layouts/15/workbench.aspx




# Deploy

1. run `gulp build`
2. run `gulp bundle --ship`
3. run `gulp package-solution --ship`
4. open `sharepoint` folder (can be seen now in VS), goto `solution` folder, and find `.sppkg` file
5. Goto tenant
6. Find AppCatalog (via SharePoint or via sp-admin-center)
7. click on `Distribute apps for SharePoint`
8. drag `.sppkg` file
9. checkbox according to need
10. if not 1st time, do check in
11. if stuck refresh page, make sure that file is checked in



gulp build; gulp bundle --ship; gulp package-solution --ship


# some debug

* npm ls -g --depth=0 @microsoft/generator-sharepoint
* node -v; gulp -v;
