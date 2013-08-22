# Esse arquivo define as configurações usadas nos arquivos de build em ./jakelib

# A solução a ser compilada
global.solutionFile = 'src/Vtex.Practices.Datatransformation.sln'

# Caminho para o arquivo de resultados gerado pelo NUnit
global.nUnitOutputFile = 'test-output/VTEX Practices Datatransformation.xml'

# Projetos de teste
global.testProjects = [
    {name: 'Vtex.Practices.Datatransformation.Tests'}    
]

# Define caminho dos assemblies dos projetos de testes
for testProject in global.testProjects
    testProject.assembly = "src/#{testProject.name}/bin/Release/#{testProject.name}.dll"

# Projetos a terem pacotes NuGet gerados
global.projects =
    main:
        name: 'Vtex.Practices.Datatransformation'
        id: 'VTEXPractices.Datatransformation'

# Define caminhos para Nuspec e Assembly Info de cada projeto
for key, project of global.projects
    project.assemblyInfo = "src/#{project.name}/Properties/AssemblyInfo.cs"
    project.nuspec = "src/#{project.name}/#{project.id}.nuspec"