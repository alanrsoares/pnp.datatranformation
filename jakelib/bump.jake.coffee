clc = require './lib/color'
require '../jakelib/lib/options'
require 'shelljs/global'

gitTagPrefix = global.options.git.tagPrefix
projects = global.projects

utils = require '../jakelib/lib/utils'
semVer = require '../jakelib/lib/semVer'
minimatch = require 'minimatch'
path = require 'path'
{AssemblyInfo} = require '../jakelib/lib/assemblyInfo'

desc \
    """
    Atualiza a versão do AssemblyInfo de um projeto  

        #{clc.title 'USO:'}
          
            #{clc.topics '$ jake bump[<incremento>,<nomedoprojeto>] prerelease=<preRelease>'}


        #{clc.title 'DESCRIÇÃO:'}

            Atualiza a versão do AssemblyInfo de um projeto, fazendo commit das
            alterações e aplicando a tag da versão no git. É possível inserir o valor de 
            prerelease na versão, bem como removê-lo quando julgar necessário. 

            O número de versão respeita as regras sugeridas pelo Semantic Versioning,portanto,
            é necessário que o usuário esteja atento aos padrões apresentados em semver.org.

            Enquanto o working directory possuir modificações locais não resolvidas, a tarefa não
            será executada. O usuário pode específicar quais arquivos podem ser commitados com a
            tag configurando o campo \'options.git.allowInBumpCommit\' no Jakefile.


        #{clc.title 'PARÂMETROS:'}
        
           #{clc.topics '<incremento> (opcional)'}
                Campo que deve ser incrementado no número de versão MAJOR.MINOR.PATCH
                do projeto, obedecendo as regras sugeridas pelo Semantic Versioning.
                Caso seja o único valor passado como parâmetro, o projeto \'default\'
                será incrementado.

            #{clc.topics '<nomedoprojeto> (opcional)'}
                Nome do projeto que se pretende incrementar o número de versão, os projetos
                devem ser configurados previamente no arquivo Jakefile.coffee.

            #{clc.topics '<preRelease> (opcional)'}
                Valor da prerelease que será adicionada a versão. Caso o valor seja vazio e a 
                versão atual já tenha uma prerelease, a prerelease será removida. É possível adicionar uma 
                prelease passando o projeto desejado como parâmetro, neste caso, valores não são
                incrementados.


         #{clc.title 'EXEMPLOS:'}

            Incrementa a versão do projeto \'default\' configurado no Jakefile.coffee:
                 
                 #{clc.example '$ jake bump[major]'}
                 #{clc.example '$ jake bump[minor]'}
                 #{clc.example '$ jake bump[patch]'}
            
            Incrementa a versão do projeto \'client\' configurado no Jakefile.coffee:

                 #{clc.example '$ jake bump[minor,client]'}

            Incrementa a versão do projeto \'default\' configurado no Jakefile.coffee,
            acrescentando um prerelease a versão:

                 #{clc.example '$ jake bump[major] prerelease=gama'}
                 #{clc.example '$ jake bump[minor] prerelease=beta'}
                 #{clc.example '$ jake bump[patch] prerelease=alfa'}

            Incrementa a versão do projeto \'client\' configurado no Jakefile.coffee, 
            acrescentando um prerelease a versão:

                 #{clc.example '$ jake bump[minor,client] prerelease=beta'}

            Insere o valor da prerelease na versão atual do projeto \'default\' sem incrementar valores:

                 #{clc.example '$ jake bump prerelease=beta'}

            Insere o valor da prerelease na versão atual do projeto \'client\' sem incrementar valores:

                 #{clc.example  '$ jake bump[client] prerelease=beta'}

            Retira o valor da prerelease na versão atual sem incrementar valores:

                 #{clc.example '$ jake bump prerelease='}


    """
task 'bump', async: true, (increment, name) ->
    ifWorkingDirectoryIsClean()
    preRelease = getPreReleaseFromCli()
    name = resolveProjectName increment, name, preRelease

    version = bumpVersion increment, name, preRelease
    tag = getTag name, version

    projectIdentifier = getProjectIdentifier name
    console.log "#{clc.message '--> Atualizando versão do '+projectIdentifier+''}"
    console.log "#{clc.note 'Versão gerada:' + tag}"
    console.log "\n#{clc.message '--> Consolidando alteração de versão no git'}"
    exec "git commit --all -m\"Gera nova versão do #{projectIdentifier}\"", (code) ->
        unless code is 0
            fail 'O commit falhou'

        jake.exec ["git tag #{tag} -f"], ->
            complete(console.log "\n #{clc.sucess '--> Nova versão gerada com sucesso.'}")

# Auxiliary functions
bumpVersion = (increment, name, preRelease) ->
    project = utils.getProject name
    version = getNextVersion project, increment, preRelease
    return version

getNextVersion = (project, increment, preRelease) ->
    assemblyInfo = utils.loadAssemblyInfo project

    if preRelease? and (semVer.isPreReleaseValid preRelease) is true
        if (semVer.isIncrementValid increment) is false
            assemblyInfo.setPreRelease preRelease
            return assemblyInfo.parseVersion().ToString()

    validatePreRelease preRelease
    return assemblyInfo.getNextVersion increment, preRelease

getTag = (project, version) ->  
    if !!gitTagPrefix
        if project is "default" then newVersion = "#{gitTagPrefix}-v#{version}"
        else newVersion = "#{gitTagPrefix}-#{project}-v#{version}"
    else
        if project is "default" then newVersion = "v#{version}"
        else newVersion = "#{project}-v#{version}"

    newVersion

getProjectIdentifier = (name) ->
    project = utils.getProject name
    if project.name isnt undefined and project.name isnt '' then return project.name
    else
        if project.id isnt undefined and project.id isnt '' then return project.id
        else
            utils.invalidValue '\'name\' ou \'id\' do projeto \''+name+'\''

ifWorkingDirectoryIsClean = ->
    result = exec('git status -s', {silent:true}).output
    result = result.split '\n'

    if result.length is 0 then return

    for item in result
        item = path.basename item
        item = item.match /([^\s]+)$/
        if item isnt null
            if (ifAllowInBumpCommit item) is false then fail \
            """
            #{clc.error 'Você possui mudanças locais não resolvidas.\n
            Certifique-se que o working directory esteja vazio.\n'}
            """

ifAllowInBumpCommit = (item) ->
    allow = options.git.allowInBumpCommit
    if allow.length is 0 then utils.invalidValue 'allowInBumpCommit'

    for allowItem in allow
        regex = getRegex allowItem
        if (regex.test item) is true then return true

    return false

getRegex = (allowItem) ->
    regex = minimatch.makeRe allowItem
    if regex is false
        fail \
        """
        #{clc.error 'O wildcard \''+allowItem+'\' é
         inválido em git.allowInBumpCommit.'}
        """
    return regex

validatePreRelease = (preRelease) ->
    if (semVer.isPreReleaseValid preRelease) is false
        msg =
        """
        #{clc.error 'O valor \''+preRelease+'\' para o prerelease é inválido.
         Verifique as regras em \'semver.org\''}
        """
        fail msg

resolveProjectName = (increment, name, preRelease) ->
    ifParamsNotUndefined increment, name, preRelease

    if (semVer.isIncrementValid increment) isnt true
        if (projects[increment]) isnt undefined
            name = increment
        else
            name = undefined

    if increment is undefined and name is undefined
        if preRelease? then name = 'default'
    
    if (semVer.isIncrementValid increment) is true and name is undefined
        name = 'default'

    return name

ifParamsNotUndefined = (increment, name, preRelease) ->
    if increment is undefined and name is undefined and preRelease is undefined
        fail \
        """
        #{clc.error '--> O projeto não pode ser incrementado.\n
        Você pode ter esquecido de passar algum parâmetro.\n
        Consulte jake help[bump]'}\n
        """

getPreReleaseFromCli = ->
    if process.env.prerelease? then return process.env.prerelease
    if process.env.pre? then return process.env.pre