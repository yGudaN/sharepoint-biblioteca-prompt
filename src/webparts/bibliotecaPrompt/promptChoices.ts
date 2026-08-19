export interface IPromptChoices {
  acoes: string[];
  segmentos: string[];
  categorias: string[];
  funcionaCom: string[];
}

export const DEFAULT_PROMPT_CHOICES: IPromptChoices = {
  acoes: ['Analisar', 'Perguntar', 'Resumir', 'Criar', 'Encontrar', 'Aprender', 'Otimizar', 'Se preparar', 'Entender'],
  segmentos: ['Comercial', 'DP', 'Financeiro', 'Infra', 'Projetos', 'RH', 'Analista Funcional', 'Desenvolvedor(a)', 'Gerente de Projetos'],
  categorias: ['Área', 'Função'],
  funcionaCom: ['Outlook', 'Teams', 'OneNote', 'Word', 'Excel', 'PowerPoint', 'Power BI', 'M365 Copilot', 'Copilot Studio', 'D365 CCaaS / Customer Service', 'D365 Customer Insights - Journeys', 'D365 Sales', 'Fabric', 'Power Apps', 'Power Automate', 'Power Pages', 'Whiteboard']
};
