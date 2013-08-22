solutionFile = global.solutionFile
nUnitOutputFile = global.nUnitOutputFile
testProjects = global.testProjects
projects = global.projects

require 'shelljs/global'
njake = require './njake/njake.js'
msbuild = njake.msbuild
nunit = njake.nunit

###
@task build
Compila a solução definida na configuração 'solutionFile'
###
desc 'Compila a solução'
task 'build', async: true, ->
    console.log "\n --> Compilando \"#{solutionFile}\""
    msbuild {
        file: solutionFile
        targets: ['Clean', 'Build']
        }, -> complete()

directory 'test-output'

###
@task test
Executa todos os testes na configuração 'testProjects'.
É utilizado o campo 'assembly' de cada projeto de teste.
###
desc 'Executa os testes automatizados'
task 'test', ['build', 'test-output'], async: true, ->
    console.log '\n --> Executando testes'
    nunit {
        assemblies: testProjects.map (testProject) -> testProject.assembly 
        xml: nUnitOutputFile
        }, -> complete()

###
@task bump[name,increment]
Atualiza a versão do AssemblyInfo de um projeto, fazendo commit das
# alterações e aplicando a tag da versão
@param name - chave do projeto a ter sua versão incrementada
@param increment - parte da versão que deve ser incrementada.
                   Pode ser "major", "minor" ou "patch".
###
desc 'Incrementa a versão'
task 'bump', async: true, (name, increment) ->
    project = getProject name
    console.log " --> Incrementando versão #{increment} do projeto #{project.name}"
    version = bumpVersion project, increment

    console.log "\n --> Consolidando alteração de versão no git"
    exec "git commit --all -m\"Gera nova versão do #{project.id}\"", (code) ->
        unless code is 0
            fail 'O commit falhou'

        jake.exec ["git tag #{name}-v#{version} -a -f"], interactive: true, ->
            complete()

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

# Auxiliary functions
getProject = (name) ->
    project = projects[name]
    unless project
        fail "O projeto #{name} não é conhecido."
    project

bumpVersion = (project, increment) ->
    data = cat project.assemblyInfo
    regex = /Version\s*\("(\d+)\.(\d+)\.(\d+)"\)/g;
    match = regex.exec(data);

    version = 
        major: match[1];
        minor: match[2];
        patch: match[3];

    oldVersion = "#{version.major}.#{version.minor}.#{version.patch}"
    version[increment]++
    if increment is 'major' then version.minor = version.patch = 0
    if increment is 'minor' then version.patch = 0
    newVersion = "#{version.major}.#{version.minor}.#{version.patch}"

    data = data.replace(regex, "Version(\"#{newVersion}\")")
    data.to project.assemblyInfo

    console.log "Versão alterada de #{oldVersion} para #{newVersion}"
    newVersion
