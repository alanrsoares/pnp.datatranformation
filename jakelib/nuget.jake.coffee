require '../jakelib/lib/options'
projects = global.projects
backupOldPackages = global.options.backupOldPackages

require 'shelljs/global'
clc = require './lib/color'
utils = require '../jakelib/lib/utils'
fs = require('fs')
path = require('path')
njake = require '../tools/njake/njake.js'
nuget = njake.nuget

{AssemblyInfo} = require '../jakelib/lib/assemblyInfo'

# NJake defaults
nuget.setDefaults
    _exe: 'tools/nuget/NuGet.exe'
    verbose: false

namespace 'nuget', ->
    desc \
        """
        Publica o pacote NuGet no MyGet
                
           #{clc.title 'USO:'}

                #{clc.topics '$ jake nuget:publish[<nomedoprojeto>] [<listadeprojetos>]'}


            #{clc.title 'DESCRIÇÃO:'}

                Publica o pacote do projeto no MyGet da VTEX, o projeto
                deve previamente ter sido compilado e empacotado.

                É possível passar um projeto específico ou uma lista de projetos
                que se deseja publicar.

                Caso não sejam passados parâmentros o projeto default será publicado. A configuração de um projeto
                no Jakefile permite dois tipos de abstração.

                É possível compor vários projetos em um único projeto,desta forma, quando o 'meta-projeto' for chamado
                irá executar todos os projetos que possui.

                Por outro lado, o projeto pode ser configurado de maneira simples como um projeto específico, sem correlações
                com outros projetos.


            #{clc.title 'PARÂMETROS:'}

                #{clc.topics '<nomedoprojeto> (opcional)'}
                    Nome do projeto a ser publicado.

                #{clc.topics '<listadeprojetos> (opcional)'}
                    Lista de projetos a serem publicados.


            #{clc.title 'EXEMPLOS:'}

                Publica pacote do projeto \'default\':

                    #{clc.example '$ jake nuget:publish'}

                Publica pacote do projeto \'client\':

                    #{clc.example '$ jake nuget:publish[client]'}

                Publica pacotes dos projetos listados:

                    #{clc.example '$ jake nuget:publish[client,server]'}


        """
    task 'publish', ['build'], async: true, ->
        if arguments.length is 0
            releaseProject 'default', 'publish'
        else
            for projectName in arguments
               releaseProject projectName, 'publish'

    desc \
        """
        Cria o pacote NuGet
                
             #{clc.title 'USO:'}

                #{clc.topics '$ jake nuget:pack[<nomedoprojeto>] [<listadeprojetos>]'}


            #{clc.title 'DESCRIÇÃO:'}

                Cria o pacote do projeto passado como parâmentro,o projeto
                deve ter sido previamente compilado.

                É possível passar um projeto específico ou uma lista de projetos
                que se deseja empacotar. Lembrando que a versão do pacote, deve
                ser gerada previamente.

                Caso não sejam passados parâmentros o projeto default será publicado. A configuração de um projeto
                no Jakefile permite dois tipos de abstração.

                É possível compor vários projetos em um único projeto,desta forma, quando o 'meta-projeto' for chamado
                irá executar todos os projetos que possui.

                Por outro lado, o projeto pode ser configurado de maneira simples como um projeto específico, sem correlações
                com outros projetos.

                Se o usuário quiser realizar backup dos pacotes gerados anteriormente,
                é necessário que configure a opção \'oldPackages\' como \'true\' através do namespace \'options\' no Jakefile.
                Por default os pacotes antigos são deletados.


            #{clc.title 'PARÂMETROS:'}

                #{clc.topics '<nomedoprojeto> (opcional)'}
                    Nome do projeto a ser empacotado.

                #{clc.topics '<listadeprojetos> (opcional)'}
                    Lista de projetos a serem empacotados.


            #{clc.title 'EXEMPLOS:'}

                Cria pacote do projeto \'default\':

                    #{clc.example '$ jake nuget:pack'}

                Cria pacote do projeto \'client\':

                    #{clc.example '$ jake nuget:publish[client]'}

                Cria pacotes dos projetos listados:

                   #{clc.example '$ jake nuget:publish[client,server]'}


        """
    task 'pack',['build',  'nuget'], async: true, ->
        if arguments.length is 0
            releaseProject 'default', 'pack'
        else
            for projectName in arguments
                releaseProject projectName, 'pack'
        complete()

directory 'nuget'

# Auxiliary functions
releaseProject = (name, option) ->
    project = utils.getProject name
    version = utils.getVersion project
    if project.projects is undefined
        if option is 'pack' then pack project, version
        if option is 'publish' then publish project, version
    else
        releaseMetaProject project, name, option, version

releaseMetaProject = (project, name, option, version) ->
    for index, project of projects[name].projects
        if option is 'pack' then pack project, version
        if option is 'publish' then publish project, version

pack = (project, version, callback) ->
    ifParamsIsValid project
    utils.createDirectory 'nuget'
    if backupOldPackages is 'true' then movePreviousPackage project.id
    else
        deletePreviousPackage project.id
    
    nuget.pack {
        nuspec: project.nuspec
        version: version.ToString()
        outputDirectory: 'nuget'
    }, callback
    console.log "\n #{clc.sucess '--> Pacote do projeto \''+project.id+'\' gerado com sucesso.'}\n"

publish = (project, version) ->
    pack project, version, ->
        packageFile = getLatestPackage project.id
        console.log "\n #{clc.message '--> Publicando pacote '+packageFile}"
        nuget.push {
            package: packageFile
            source: 'https://www.myget.org/F/vtexlab/'
        }, -> complete(console.log "\n #{clc.sucess '--> Pacote '+packageFile+' publicado com sucesso.'}")

getLatestPackage = (id) ->
    all = getPackages id
    unless all.length
        msg = \
        """
        #{clc.error 'O projeto \''+id+'\' não possui um pacote correspondente.\n'}
        """
        throw new Error msg

    all[all.length - 1].replace /\//g, '\\'

deletePreviousPackage = (id) ->
    all = getPackages id
    
    console.log "\n #{clc.message '--> Removendo pacotes antigos'}"
    for file in all
        if fs.statSync(file).isFile() then fs.unlinkSync(file)
  
    console.log "\n #{clc.sucess '--> Pacotes removidos com sucesso.'}"

movePreviousPackage = (id) ->
    all = getPackages id
    backupOldPackagesPath = "nuget/backup/"
    utils.createDirectory backupOldPackagesPath
    
    console.log "\n #{clc.message '--> Realizando backup dos pacotes antigos'}"
    for file in all
        fileName = path.basename file
        try 
            fs.renameSync(file, backupOldPackagesPath+fileName)
        catch error
            console.error error

    console.log "\n #{clc.message '--> Backup realizado com sucesso.'}"

getPackages = (id) ->
    regex = new RegExp "#{id}\\.((\\d+)\.(\\d+)\\.(\\d+))(.|-([a-z]+).)nupkg"
    all = ls('nuget/*.nupkg')
    files = []

    for file in all
        match = regex.test file
        if match isnt false then files.push file.toString()
          
    return files

ifParamsIsValid = (project) ->
    if project.nuspec is undefined or project.nuspec is ''
        utils.invalidValue 'nuspec'
    if project.id is undefined or project.id is ''
        utils.invalidValue 'id'