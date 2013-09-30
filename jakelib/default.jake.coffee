clc = require './lib/color'
require '../jakelib/lib/options'

nUnitOutputFile = global.options.nUnit.outputFile
solutionFile = global.solutionFile
testProjects = global.testProjects || []
projects = global.projects

require 'shelljs/global'
njake = require '../tools/njake/njake.js'
Table = require 'cli-table'
utils = require '../jakelib/lib/utils'

msbuild = njake.msbuild
nunit = njake.nunit

desc \
    """
    Compila a solução definida na configuração \'solutionFile\'
            
        #{clc.title 'USO:'}
          
            #{clc.topics '$ jake build'}


        #{clc.title 'DESCRIÇÃO:'}

            Compila a solução previamente configurada no arquivo Jakefile.coffee 
            no campo \'solutionFile\', que deve representar o caminho para o diretório
            da solução do projeto.

            Não são necessários parâmetros para executar a tarefa, basta apenas que o campo \'solutionFile\'
            tenha sido configurado, o resultado da compilação pode ser avaliado na pasta \'/bin\'
            do projeto.


        #{clc.title 'PARÂMETROS:'}

             #{clc.topics 'n/a'}


        #{clc.title 'EXEMPLOS:'}

            Compila a solução configurada:

                #{clc.example '$ jake build'}


    """
task 'build', async: true, ->
    if not solutionFile then utils.invalidValue 'solutionFile'

    console.log "\n #{clc.message '--> Compilando '+'\"'+solutionFile+'\"'}"
    msbuild {
        file: solutionFile
        targets: ['Clean', 'Build']
        }, -> complete(console.log "\n #{clc.sucess '--> \"'+solutionFile+'\" foi compilado com sucesso.'}")

directory 'test-output'

desc \
    """
    Executa os testes automatizados
            
        #{clc.title 'USO:'}
          
            #{clc.topics '$ jake test'}


        #{clc.title 'DESCRIÇÃO:'}

            Executa todos os testes previamente configurados no arquivo Jakefile.coffee
            no campo \'testProjects\', responsável por identificar quais são os projetos
            de testes que deverão ser executados.

            Para imprimir os resultados dos testes em um arquivo, é preciso configurar o campo \'outputFile\' 
            através do namespace \'options\' no Jakefile.

            Não são necessários parâmetros para executar a tarefa, basta apenas que o campo \'testProjects\'
            tenha sido configurado.


        #{clc.title 'PARÂMETROS:'}
        
           #{clc.topics 'n/a'}


        #{clc.title 'EXEMPLOS:'}

            Testa projetos configurados:

                 #{clc.example '$ jake test'}


    """
task 'test', ['build', 'test-output'], async: true, ->
    if testProjects.length is 0 then utils.invalidValue 'testProjects'

    console.log "\n #{clc.message "--> Executando testes"}"
    nunit {
        assemblies: testProjects.map (testProject) -> testProject.assembly 
        xml: nUnitOutputFile
        }, -> complete(console.log "\n #{clc.sucess '--> Testes executados com sucesso.'}")

# NJake defaults
msbuild.setDefaults
    properties:
        configuration: 'Release'
        warningLevel: 1
    processor: 'x86'
    version: 'net4.0'
    _parameters: ['/verbosity:minimal']

nunit.setDefaults
    _exe: 'tools/nunit/nunit-console-x86.exe'

desc \
    """
    Lista os projetos configurados em Jakefile.coffee
            
        #{clc.title 'USO:'}
          
            #{clc.topics '$ jake projects'}


       #{clc.title 'DESCRIÇÃO:'}

            Lista todos os projetos configurados no Jakefile.coffee, os projetos
            são apresentados em uma tabela com as seguintes colunas: \'alias\', \'name\'' e \'id\' do projeto.


       #{clc.title 'PARÂMETROS:'}
        
           #{clc.topics 'n/a'}


        #{clc.title 'EXEMPLOS:'}

            Lista todos os projetos:

                #{clc.example '$ jake projects'}


    """
task 'projects',  ->
    table = new Table({head:['alias','name','id']})

    for key, value of projects
        if value.name is undefined then value.name = 'não definido'
        if value.id is undefined then value.id = 'não definido'
        table.push([key, value.name, value.id])

    console.log(table.toString())

desc \
    """
    Atualiza a versão do VTEX Jake Bootstrapper

       #{clc.title 'USO:'}

            #{clc.topics '$ jake update'}


       #{clc.title 'DESCRIÇÃO:'}

            Atualiza a ferramenta VTEX Jake Bootstrapper para sua versão mais recente.


        #{clc.title 'PARÂMETROS:'}

            #{clc.topics 'n/a'}


        #{clc.title 'EXEMPLOS:'}

            Atualiza a versão do VTEX Jake Bootstrapper:

                 #{clc.example '$ jake update'}


    """
task 'update', async: true, ->
    packageJSON = require '../node_modules/jake-bootstrapper/package.json'
    if packageJSON._from?
        jake.exec "npm install #{packageJSON._from}", {interactive:true}, -> complete()
    else
        console.log "#{clc.error 'O arquivo package.json não foi encontrado.'}"
        complete()

desc \
    """
    Remove arquivos dos diretórios 'obj' e 'bin'

        #{clc.title 'USO:'}

            #{clc.topics '$ jake clean'}


        #{clc.title 'DESCRIÇÃO:'}

            Remove arquivos dos diretórios 'obj'e 'bin' do projeto.


        #{clc.title 'PARÂMETROS:'}

                #{clc.topics 'n/a'}


        #{clc.title 'EXEMPLOS:'}

            Remove arquivos do diretórios 'bin' e 'obj' do projeto 'default':

                #{clc.example '$ jake clean'}


    """
task 'clean', ->
    path = solutionFile.replace /\/([\w\s.]+).sln$/, ''
    listBinObj = findBinObjDirectory path
    cleanDirectory listBinObj

# Auxiliary functions
findBinObjDirectory = (path) ->
    list = ls '-R', path
    if list.length is 0
        fail "#{clc.error 'O diretório \''+path+'\' não foi encontrado.'}"

    listBinDir = find('.').filter((list) -> /bin$/.exec list)
    listObjDir = find('.').filter((list) -> /obj$/.exec list)
    listBinObj = listBinDir.concat listObjDir
    return listBinObj

cleanDirectory = (listBinObj) ->
    for dir in listBinObj
        if (test '-d', dir)
            rm '-rf', dir+'/*'
        else
            console.log fail "#{clc.error '--> Erro ao remover os arquivos.\n'}"

    console.log "\n#{clc.sucess '--> Arquivos removidos com sucesso.'}\n"