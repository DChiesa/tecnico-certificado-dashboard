
let supabaseClient;

function configurarSupabase() {
  if (!window.APP_CONFIG) {
    throw new Error(
      "config.js não foi carregado."
    );
  }

  if (!window.supabase) {
    throw new Error(
      "Biblioteca Supabase não foi carregada."
    );
  }

  supabaseClient = window.supabase.createClient(
    window.APP_CONFIG.SUPABASE_URL,
    window.APP_CONFIG.SUPABASE_PUBLISHABLE_KEY,
    {
      auth: {
        persistSession: true,
        autoRefreshToken: true,
        detectSessionInUrl: true
      }
    }
  );
}

async function obterSessaoSupabase() {
  if (!supabaseClient) {
    configurarSupabase();
  }

  const {
    data: { session },
    error
  } = await supabaseClient.auth.getSession();

  if (error) {
    throw error;
  }

  return session;
}

async function entrarNoDashboard(email, senha) {
  if (!supabaseClient) {
    configurarSupabase();
  }

  const { data, error } =
    await supabaseClient.auth.signInWithPassword({
      email: email,
      password: senha
    });

  if (error) {
    throw error;
  }

  return data.session;
}

async function sairDoDashboard() {
  if (!supabaseClient) {
    configurarSupabase();
  }

  await supabaseClient.auth.signOut();
  window.location.reload();
}

function mostrarTelaLogin() {
  const tela =
    document.getElementById("loginScreen");

  if (tela) {
    tela.style.display = "flex";
  }
}

function esconderTelaLogin() {
  const tela =
    document.getElementById("loginScreen");

  if (tela) {
    tela.style.display = "none";
  }
}

async function baixarExcelSupabase(
  caminho,
  nomeExibido
) {
  if (!supabaseClient) {
    configurarSupabase();
  }

  if (typeof setAutoStatus === "function") {
    setAutoStatus(
      "loading",
      "Baixando Excel do Supabase..."
    );
  }

  const { data, error } =
    await supabaseClient.storage
      .from(window.APP_CONFIG.BUCKET)
      .download(caminho);

  if (error) {
    throw error;
  }

  const buffer = await data.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true
  });

  if (typeof processWorkbook !== "function") {
    throw new Error(
      "processWorkbook não foi encontrada no dashboard."
    );
  }

  const total = processWorkbook(
    workbook,
    nomeExibido || caminho
  );

  const agora =
    new Date().toLocaleString("pt-BR");

  if (typeof setAutoStatus === "function") {
    setAutoStatus(
      "ok",
      `${total.toLocaleString("pt-BR")} registros · ${agora}`
    );
  }

  return total;
}

async function carregarBaseAtualSupabase() {
  try {
    const session =
      await obterSessaoSupabase();

    if (!session) {
      mostrarTelaLogin();

      if (
        typeof setAutoStatus === "function"
      ) {
        setAutoStatus(
          "error",
          "Faça login para carregar a base."
        );
      }

      return;
    }

    esconderTelaLogin();

    await baixarExcelSupabase(
      window.APP_CONFIG.ARQUIVO_ATUAL,
      "Técnico certificado atual · Supabase"
    );
  } catch (error) {
    console.error(error);

    if (
      typeof setAutoStatus === "function"
    ) {
      setAutoStatus(
        "error",
        `Falha: ${error.message}`
      );
    }
  }
}

async function listarHistoricosSupabase() {
  if (!supabaseClient) {
    configurarSupabase();
  }

  const { data, error } =
    await supabaseClient.storage
      .from(window.APP_CONFIG.BUCKET)
      .list(
        window.APP_CONFIG.PASTA_HISTORICO,
        {
          limit: 100,
          sortBy: {
            column: "name",
            order: "desc"
          }
        }
      );

  if (error) {
    throw error;
  }

  return (data || []).filter(
    item =>
      item.name
        .toLowerCase()
        .endsWith(".xlsx")
  );
}

async function preencherHistoricos() {
  const select =
    document.getElementById(
      "historicoSelect"
    );

  if (!select) {
    return;
  }

  const arquivos =
    await listarHistoricosSupabase();

  select.innerHTML =
    '<option value="">Base atual</option>';

  arquivos.forEach(arquivo => {
    const option =
      document.createElement("option");

    option.value =
      window.APP_CONFIG.PASTA_HISTORICO +
      "/" +
      arquivo.name;

    option.textContent = arquivo.name;

    select.appendChild(option);
  });
}
