clc = require './lib/color'

desc \
    """
    Exibe manual de ajuda com instruções de uso
            
        #{clc.title 'USO:'}
          
            #{clc.topics '$ jake help[<tarefa>]'}


        #{clc.title 'DESCRIÇÃO:'}

            Exibe instruções de uso da ferramenta, lista quais são as tarefas
            disponíveis apresentando uma pequena descrição de cada tarefa. 
            Caso seja passada uma tarefa como parâmetro, apresenta informações 
            detalhadas da tarefa em questão.

            A exibição detalhada de uma tarefa, identifica quais são os parâmentros
            possíveis de serem utilizados. Além disso, exemplos são fornecidos afim
            de demonstrar sua aplicabilidade.


        #{clc.title 'PARÂMETROS:'}
        
            #{clc.topics '<tarefa> (opcional)'}
                Nome da tarefa que se pretende obter informações mais detalhadas.


        #{clc.title 'EXEMPLOS:'}

            Acesso ao menu de ajuda:

                #{clc.example '$ jake help'}
            
            Acesso a informações detalhadas de uma tarefa:

                #{clc.example '$ jake help[bump]'}


    """
task 'help', (@task) ->
    regex = /^.*$/m
    if @task?
        printMsg @task
    else
      console.log "\n"
      tab = " "
      maxTaskName = maxTaskNameLength()
      for name, @task of jake.Task when @task.description?
          padding = getPadding name.length, maxTaskName
          desc = regex.exec @task.description
          msg = \
          """
            #{tab} #{clc.title name} #{padding} #{clc.comment '#', desc[0]}
          """
          list = "" + msg
          console.log list
      msg = \
      """
      
        
        Consulte #{clc.note('jake help[tarefa]')} para informações específicas sobre uma tarefa.
        Para realizar as configurações, altere o #{clc.note('arquivo Jakefile')}.

      """        
      console.log msg

task 'default', ['help'], -> complete()

# Auxiliary functions
getPadding = (length, maxTaskName) ->
    padding = (new Array(maxTaskName - length + 2)).join(' ');
    return padding

maxTaskNameLength = () ->
  maxTaskNameLength = 0
  for index, @task of jake.Task
     if index.length > maxTaskNameLength
        maxTaskNameLength = index.length
  
  return maxTaskNameLength

printMsg = (@task) ->
    regex = /^.*$/m
    for name, @taskName of jake.Task when @taskName.description?
        if @task is name
            console.log @taskName.description.replace regex, ' '
            return

    console.log "\n #{clc.error 'A tarefa \''+@task+'\' não existe. Consulte \'jake help\''}."
