projects = global.projects

require 'shelljs/global'
njake = require './njake/njake.js'
nuget = njake.nuget

# NJake defaults
nuget.setDefaults
    _exe: 'tools/nuget/NuGet.exe'
    verbose: false

namespace 'nuget', ->
    ###
    @task publish
    Gera pacote NuGet e publica no MyGet da VTEX
    @param projectName - chave do projeto a ser publicado
    ###
    desc 'Publica o pacote NuGet no MyGet'
    task 'publish', ['build'], async: true, (projectName) ->
        project = projects[projectName]
        pack project, ->
            packageFile = getLatestPackage project.id
            console.log "\n --> Publicando pacote #{packageFile}"
            nuget.push {
                package: packageFile
                source: 'https://www.myget.org/F/vtexlab/'
            }, -> complete()

# Auxiliary functions
pack = (project, callback) ->
    console.log "\n --> Criando pacote NuGet para #{project.name}"
    nuget.pack {
        nuspec: project.nuspec
        version: getAssemblyVersion project.assemblyInfo
        outputDirectory: 'nuget'
    }, callback

getAssemblyVersion = (assemblyInfoFile) ->
    data = cat assemblyInfoFile
    regex = /Version\s*\("(\d+\.\d+\.\d+)"\)/g
    match = regex.exec(data)
    match[1]

getLatestPackage = (id) ->
    regex = new RegExp "#{id}\\.[\\.\\d]+\\.nupkg$"
    all = find('nuget').filter (file) -> regex.test file
    all[all.length - 1].replace /\//g, '\\'
