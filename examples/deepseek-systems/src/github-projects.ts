const repositories = ["FlashMLA", "DeepEP", "DeepGEMM"] as const;

export async function fetchGitHubProjects(fetchProject: typeof fetch = fetch) {
  const projects = await Promise.all(
    repositories.map(async (name) => {
      const repository = `deepseek-ai/${name}`;
      const signal = AbortSignal.timeout(10_000);
      try {
        const response = await fetchProject(
          `https://api.github.com/repos/${repository}`,
          { signal },
        );
        if (!response.ok) {
          throw new Error(`GitHub returned HTTP ${response.status}`);
        }
        const data: unknown = await response.json();
        if (
          typeof data !== "object" ||
          data === null ||
          !("name" in data) ||
          typeof data.name !== "string" ||
          data.name !== name ||
          !("html_url" in data) ||
          typeof data.html_url !== "string" ||
          data.html_url !== `https://github.com/${repository}` ||
          !("stargazers_count" in data) ||
          typeof data.stargazers_count !== "number" ||
          !Number.isSafeInteger(data.stargazers_count) ||
          data.stargazers_count < 0
        ) {
          throw new Error("Invalid repository data");
        }
        return {
          name: data.name,
          url: data.html_url,
          stars: data.stargazers_count,
        };
      } catch (cause) {
        const reason = signal.aborted
          ? "Request timed out after 10 seconds"
          : cause instanceof Error
            ? cause.message
            : String(cause);
        throw new Error(`Failed to fetch ${repository}: ${reason}`, { cause });
      }
    }),
  );

  return { projects, fetchedAt: new Date().toISOString() };
}
