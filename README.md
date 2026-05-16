# Net-Diagnostic

🚀 Funcionalidades

* 🔍 Diagnóstico inteligente automático
    * Consolida todos os testes e identifica padrões de falha
    * Classificação por severidade: OK, ALERTA, CRÍTICO
* 🌐 Testes de conectividade
    * Ping (latência e perda)
    * TCP (porta aberta/fechada)
    * DNS (resolução e tempo de resposta)
* 🧭 TraceRoute automatizado
    * Executado sob condições específicas:
        * Falha de conectividade externa
        * Latência elevada
        * Perda de pacotes
    * Geração de resumo e análise de rota
* 📊 Monitoramento persistente
    * Execução contínua com coleta de eventos
    * Registro histórico de falhas intermitentes
* 🧠 Sistema de scoring
    * Avaliação da saúde da rede
    * Pontuação baseada em múltiplos fatores
* 📁 Geração de relatórios
    * Estruturado por seções:
        * Resumo geral
        * Diagnóstico inteligente
        * Testes executados
        * Eventos detectados
        * Tracert (resumo e completo)




⚙️ Lógica de Diagnóstico

O sistema aplica heurísticas para identificar a origem do problema:

* Gateway OK + Internet falhando → problema externo
* Alta latência (>150ms) → possível congestionamento
* Perda de pacotes → instabilidade de rede
* DNS lento ou falhando → possível gargalo de resolução

Quando detectado:

➡️ Executa automaticamente TraceRoute para análise detalhada da rota
