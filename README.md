# speedPC

**Ferramenta gratuita de diagnóstico e manutenção de computadores Windows para empresas.**

O speedPC ajuda o técnico de TI a descobrir **por que** um computador está com problema, e não só a "dar uma limpeza". Ele observa a máquina e lê o histórico do próprio Windows para apontar **o problema, a evidência e o que fazer**:

- **Lentidão:** PC sem reiniciar há dias, memória RAM esgotada (e quem está consumindo), processador superaquecendo, disco cheio ou HD antigo, inicialização lenta e os programas que mais atrasam.
- **Travamentos:** telas azuis com o código traduzido para a causa, desligamentos inesperados (inclusive quando o PC **trava ao desligar**), erros de hardware e de disco.
- **Câmera e microfone:** desativados, bloqueados na privacidade ou por política, ou presos em outro programa (o clássico "a câmera não abre no Teams porque o Zoom está aberto").
- **Rede e impressão:** sem IP, DNS falhando, Wi-Fi fraco, cabo limitado a 100 Mbps, fila de impressão travada.

No final ele oferece as **correções automáticas possíveis**, e o técnico escolhe quais aplicar. Também faz **limpeza segura** de todos os usuários, **otimização**, **reparo do Windows**, **atualizações e drivers**, **instalação de aplicativos** pelo winget, gera um **relatório HTML por máquina** para anexar no chamado e pode rodar **agendado toda semana**, salvando os relatórios numa pasta de rede.

**Seguro e reversível:** ponto de restauração, backup de cada configuração alterada e opção de "Desfazer". É **PowerShell puro**, num único arquivo (`speedPC.ps1`) fácil de auditar, sem programas de terceiros. Funciona no Windows 10 e 11 com o PowerShell 5.1 que já vem no Windows.

## Como usar

1. Baixe o projeto (botão **Code › Download ZIP** aqui no GitHub) e extraia a pasta no computador, num pendrive ou numa pasta de rede. Se o Windows bloquear o arquivo baixado, clique com o botão direito no `.zip`, abra **Propriedades** e marque **Desbloquear** antes de extrair.
2. Dê dois cliques em **`Executar_speedPC.bat`**. O script pede permissão de administrador sozinho e abre a **janela gráfica** (WPF, que já vem no Windows).
3. A aba **Visão geral** mostra o essencial:
   - **há quanto tempo o PC está ligado** (contando do último boot real);
   - **quando e como foi o último desligamento**: normal (quem pediu e por qual programa) ou inesperado (tela azul, botão de energia segurado, falta de energia, travou ao desligar);
   - **uso da memória RAM** e os programas que mais consomem, espaço em disco e a ficha da máquina.

Prefere o terminal? Use `speedPC.ps1 -Console` para o menu em modo texto.

### Abas

| Aba | O que faz |
|---|---|
| Visão geral | Resumo da máquina e ações rápidas. |
| Diagnóstico | Problemas com causa provável, evidência e o que fazer. O botão **Aplicar correções** lista as correções possíveis e o técnico escolhe quais aplicar. Gera o relatório HTML. |
| Limpeza | Temporários e caches de **todos os usuários** e do Windows. Lixeira e WinSxS são opcionais. |
| Otimização | Plano de energia, Inicialização Rápida, sugestões e anúncios, TRIM/desfragmentação, Sensor de Armazenamento, programas de inicialização e o plano Desempenho Máximo (desktops). |
| Reparo | DISM + SFC, CHKDSK online, reparo do Windows Update, renovação de IP/DNS, reset de rede, fila de impressão, teste de memória e atalhos para os painéis clássicos do Windows. |
| Aplicativos | Instala, atualiza ou desinstala programas pelo **winget** (repositório oficial da Microsoft): navegadores, comunicação, PDF e escritório, acesso remoto, utilitários, diagnóstico de hardware, rede e TI, senhas e componentes (Visual C++, .NET, Java). Tem busca, e os programas já instalados aparecem com a marca **INSTALADO**. Para incluir outro programa, adicione o Id em `$Script:CatalogoApps` no script (descubra o Id com `winget search nome`) e, se quiser o ícone, o site oficial em `Site`. Os ícones são baixados uma única vez, pelo serviço de favicons do Google (só o endereço do site do programa é enviado), e ficam guardados em `%LOCALAPPDATA%\speedPC\Icones`. Sem internet, cada app aparece com a inicial colorida. |
| Atualizações | Windows Update (sem prévias e sem troca de versão), drivers, programas via winget e o **perfil de atualizações para empresas** (versões novas após 365 dias e qualidade após 4 dias). |
| Automação e segurança | Agendamento semanal, desfazer sessões de alterações e ponto de restauração. |

Nas abas de opções os itens recomendados já vêm marcados. Os itens marcados com **CUIDADO** pedem atenção antes de usar. Tudo o que acontece aparece no **Registro de atividade**, na parte de baixo da janela. As ferramentas demoradas (DISM, SFC, CHKDSK, winget) abrem uma janela própria para você acompanhar o progresso.

O visual e algumas funções (painéis clássicos, perfil de atualizações, Desempenho Máximo) foram inspirados no [WinUtil do Chris Titus Tech](https://christitus.com/windows-tool/), adaptados para o uso em empresas.

## O que a análise verifica

**Lentidão:** PC sem reiniciar há dias (avisa se a Inicialização Rápida impede o reinício real), RAM esgotada e quem está consumindo, processador sobrecarregado, superaquecimento (sensor ACPI e eventos de limitação de velocidade), disco cheio, HD mecânico/antigo, saúde SMART, desgaste do SSD, tempo de boot medido pelo Windows e **os programas, drivers e serviços que mais atrasaram a inicialização**, plano de energia em economia.

**Travamentos:** telas azuis com o **código traduzido para a causa** (mais de 40 códigos catalogados), desligamentos inesperados com a causa (botão de energia, falta de energia, tela azul), erros de hardware (WHEA), erros de disco e do NTFS, e programas que fecham ou congelam com frequência.

**Câmera e microfone:** dispositivo desativado ou com erro, câmera não detectada (tecla de privacidade ou BIOS), bloqueio na privacidade (dispositivo, aplicativos, programas da área de trabalho, permissão individual do Teams), bloqueio por política (GPO/Intune), serviços de áudio parados, e **qual programa está segurando a câmera agora** (o clássico "a câmera não abre no Teams porque o Zoom está aberto"), com opção de fechar esse programa.

**Rede:** sem IP (APIPA/DHCP), sem gateway, DNS falhando, sem internet, proxy ou portal de login, Wi-Fi fraco (com banda e canal), **cabo limitado a 100 Mbps** ou half-duplex, DNS que não encontra o domínio e relação de confiança com o domínio quebrada.

**Impressão:** spooler parado, fila travada (trabalhos com erro ou parados há mais de 20 minutos), impressora padrão offline e spooler que trava por causa de driver.

**Sistema:** reinício pendente, dias sem atualização, antivírus desligado ou desatualizado, dispositivos sem driver e desgaste da bateria.

## Limpeza segura: o que entra e o que não entra

**Limpa:** `%TEMP%` dos usuários (só arquivos com mais de 2 dias), cache do Chrome/Edge/Brave/Firefox (só as pastas de cache), cache do Teams clássico, relatórios de erro, cache de miniaturas, `C:\Windows\Temp`, downloads antigos do Windows Update (não mexe se houver reinício pendente) e o cache da Otimização de Entrega. Opcional: Lixeira e limpeza de componentes (WinSxS).

**Nunca toca em:** senhas, cookies, histórico, favoritos, Downloads, Documentos, Área de Trabalho, pasta temporária de anexos do Outlook (`Content.Outlook`) nem nos despejos de tela azul (`C:\Windows\Minidump`), que servem de evidência. A limpeza não segue junções nem links simbólicos.

## Segurança e reversão

- Antes da primeira alteração da sessão é criado um **ponto de restauração** (se a Proteção do Sistema estiver desligada, a ferramenta pergunta se você quer ligar).
- Cada configuração alterada (registro, serviço, dispositivo, plano de energia, adaptador de rede) tem o **valor anterior salvo** em `C:\ProgramData\speedPC\Backups\*.json`.
- **Automação e segurança › Desfazer alterações** reverte uma sessão inteira na ordem inversa.
- Cada correção mostra se pode ser desfeita, se não muda configurações (por exemplo, a limpeza) ou se não é reversível (por exemplo, fechar um programa ou instalar atualizações).
- Políticas de GPO/Intune são apenas apontadas e nunca alteradas localmente.
- Log de tudo em `C:\ProgramData\speedPC\Logs`.

## Modo automático (semanal)

Pela aba **Automação e segurança**, ou pela linha de comando (por exemplo, numa GPO de inicialização ou pela sua ferramenta de acesso remoto):

```bash
powershell -ExecutionPolicy Bypass -File speedPC.ps1 -Agendar -PastaRelatorios \\servidor\manutencao$ -DiaSemana Friday -Hora 12:30
```

- Copia o script para `C:\Program Files\speedPC\` (só administradores podem alterá-lo lá) e cria a tarefa **\speedPC\speedPC - Manutencao semanal**, que roda como SYSTEM.
- A cada execução faz o **diagnóstico, a limpeza segura e o relatório**. Ela **não** aplica correções sozinha.
- Salva `speedPC_<PC>_<data>.html` e `<PC>_ultimo.json` (resumo para consolidar várias máquinas) na pasta de rede, além de uma cópia local em `C:\ProgramData\speedPC\Relatorios` (mantém 90 dias).
- **Permissão da pasta de rede:** como a tarefa roda como SYSTEM, quem grava é a **conta do computador**. Dê permissão de gravação ao grupo **Computadores do domínio** (Domain Computers) no compartilhamento e no NTFS.
- Para remover: `speedPC.ps1 -RemoverAgendamento`.

## Outros parâmetros

| Parâmetro | Uso |
|---|---|
| `-Console` | Abre o menu em modo texto em vez da janela gráfica. |
| `-SomenteRelatorio [-PastaRelatorios X]` | Só diagnostica e gera o relatório. Não altera nada e não pede administrador (com menos detalhes). |
| `-Auto [-PastaRelatorios X] [-SemLimpeza]` | Executa o modo automático uma vez, sem perguntas. |

## Requisitos e limites

- Windows 10 ou 11, PowerShell 5.1. A maioria das verificações e todas as correções exigem administrador.
- O sensor de temperatura ACPI não existe em todos os computadores. Nesse caso a análise usa os eventos de limitação de velocidade do processador.
- Desde o Windows 11 24H2, a leitura do sinal do Wi-Fi exige que os Serviços de Localização estejam ativos.
- Em Windows Server não há ponto de restauração, mas o backup das configurações continua funcionando.

## Licença

Distribuído sob a licença MIT: use, modifique e distribua à vontade, inclusive em empresas. Veja o arquivo [LICENSE](LICENSE).
