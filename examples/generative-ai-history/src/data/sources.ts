export type Source = {
  readonly id: number;
  readonly short: string;
  readonly title: string;
  readonly detail: string;
  readonly url: string;
};

export const sources: readonly Source[] = [
  {
    id: 1,
    short: "Turing, 1950",
    title: "Computing Machinery and Intelligence",
    detail: "Alan M. Turing · Mind 49 (236) · 1950",
    url: "https://doi.org/10.1093/mind/LIX.236.433",
  },
  {
    id: 2,
    short: "Weizenbaum, 1966",
    title: "ELIZA — A Computer Program for the Study of Natural Language Communication between Man and Machine",
    detail: "Joseph Weizenbaum · Communications of the ACM 9 (1) · 1966",
    url: "https://dl.acm.org/doi/10.1145/365153.365168",
  },
  {
    id: 3,
    short: "Kingma & Welling, 2013",
    title: "Auto-Encoding Variational Bayes",
    detail: "Diederik P. Kingma, Max Welling · arXiv:1312.6114 · 2013",
    url: "https://arxiv.org/abs/1312.6114",
  },
  {
    id: 4,
    short: "Goodfellow et al., 2014",
    title: "Generative Adversarial Nets",
    detail: "Ian Goodfellow et al. · NeurIPS 2014",
    url: "https://papers.nips.cc/paper/5423-generative-adversarial-nets",
  },
  {
    id: 5,
    short: "Vaswani et al., 2017",
    title: "Attention Is All You Need",
    detail: "Ashish Vaswani et al. · NeurIPS 2017",
    url: "https://arxiv.org/abs/1706.03762",
  },
  {
    id: 6,
    short: "Radford et al., 2018",
    title: "Improving Language Understanding by Generative Pre-Training",
    detail: "Alec Radford et al. · OpenAI · 2018",
    url: "https://cdn.openai.com/research-covers/language-unsupervised/language_understanding_paper.pdf",
  },
  {
    id: 7,
    short: "Radford et al., 2019",
    title: "Language Models are Unsupervised Multitask Learners",
    detail: "Alec Radford et al. · OpenAI · 2019",
    url: "https://cdn.openai.com/better-language-models/language_models_are_unsupervised_multitask_learners.pdf",
  },
  {
    id: 8,
    short: "Brown et al., 2020",
    title: "Language Models are Few-Shot Learners",
    detail: "Tom B. Brown et al. · NeurIPS 2020",
    url: "https://arxiv.org/abs/2005.14165",
  },
  {
    id: 9,
    short: "Ho et al., 2020",
    title: "Denoising Diffusion Probabilistic Models",
    detail: "Jonathan Ho, Ajay Jain, Pieter Abbeel · NeurIPS 2020",
    url: "https://arxiv.org/abs/2006.11239",
  },
  {
    id: 10,
    short: "Radford et al., 2021",
    title: "Learning Transferable Visual Models From Natural Language Supervision",
    detail: "Alec Radford et al. · ICML 2021",
    url: "https://arxiv.org/abs/2103.00020",
  },
  {
    id: 11,
    short: "Rombach et al., 2022",
    title: "High-Resolution Image Synthesis with Latent Diffusion Models",
    detail: "Robin Rombach et al. · CVPR 2022",
    url: "https://arxiv.org/abs/2112.10752",
  },
  {
    id: 12,
    short: "Ouyang et al., 2022",
    title: "Training Language Models to Follow Instructions with Human Feedback",
    detail: "Long Ouyang et al. · NeurIPS 2022",
    url: "https://arxiv.org/abs/2203.02155",
  },
  {
    id: 13,
    short: "OpenAI, 2022",
    title: "Introducing ChatGPT",
    detail: "OpenAI · 2022-11-30",
    url: "https://openai.com/index/chatgpt/",
  },
  {
    id: 14,
    short: "Yao et al., 2022",
    title: "ReAct: Synergizing Reasoning and Acting in Language Models",
    detail: "Shunyu Yao et al. · ICLR 2023",
    url: "https://arxiv.org/abs/2210.03629",
  },
  {
    id: 15,
    short: "OpenAI, 2023",
    title: "GPT-4 Technical Report",
    detail: "OpenAI · arXiv:2303.08774 · 2023",
    url: "https://arxiv.org/abs/2303.08774",
  },
  {
    id: 16,
    short: "OpenAI, 2024",
    title: "Hello GPT-4o",
    detail: "OpenAI · 2024-05-13",
    url: "https://openai.com/index/hello-gpt-4o/",
  },
  {
    id: 17,
    short: "OpenAI, 2024",
    title: "Learning to Reason with LLMs",
    detail: "OpenAI · 2024-09-12",
    url: "https://openai.com/index/learning-to-reason-with-llms/",
  },
  {
    id: 18,
    short: "OpenAI, 2024",
    title: "Video Generation Models as World Simulators",
    detail: "OpenAI · 2024-02-15",
    url: "https://openai.com/index/video-generation-models-as-world-simulators/",
  },
];
