
const URL_CATALOGO = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQzQhxeH16fGky6EVaWV0tQvhhvh1EEqJjWkqPmxTa1Qo0C8I0z3Wj83vsbVIVL-z3tB7cwD2t_4YE2/pub?gid=0&single=true&output=csv";

const formulario = document.getElementById("formulario-busqueda");
const campoBusqueda = document.getElementById("busqueda");
const resultados = document.getElementById("resultados");
const estado = document.getElementById("estado");

let productos = [];
let cargando = false;
let ultimaCargaCorrecta = false;

// Normaliza texto para buscar sin distinguir mayúsculas ni tildes.
function normalizar(texto) {
  return String(texto ?? "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

// Convierte una fila del CSV en un producto.
function convertirProducto(fila) {
  const datos = {};

  for (const [clave, valor] of Object.entries(fila)) {
    datos[normalizar(clave)] = valor;
  }

  function leerNumero(valor) {
    if (valor === "" || valor == null) return null;

    const numero = Number(valor);
    return Number.isFinite(numero) ? numero : null;
  }

  return {
    producto: String(datos.producto ?? "").trim(),
    categoria: String(datos.categoria ?? "").trim(),
    marca: String(datos.marca ?? "").trim(),
    precio: leerNumero(datos.precio),
    stock: leerNumero(datos.stock)
  };
}

// Carga los productos publicados en Google Sheets.
async function cargarCatalogo() {
  if (cargando) return;

  cargando = true;

  try {
    const separador = URL_CATALOGO.includes("?") ? "&" : "?";
    const url = `${URL_CATALOGO}${separador}_=${Date.now()}`;

    const respuesta = await fetch(url, { cache: "no-store" });

    if (!respuesta.ok) {
      throw new Error(`Google respondió con HTTP ${respuesta.status}`);
    }

    const csv = await respuesta.text();
    const contenido = csv.trim();

    if (
      !contenido ||
      /^<!doctype html/i.test(contenido) ||
      /^<html/i.test(contenido)
    ) {
      throw new Error("Google no devolvió un catálogo CSV válido.");
    }

    const libro = XLSX.read(csv, { type: "string" });
    const hoja = libro.Sheets[libro.SheetNames[0]];

    if (!hoja) {
      throw new Error("No se encontró la hoja del catálogo.");
    }

    const filas = XLSX.utils.sheet_to_json(hoja, {
      defval: "",
      raw: false
    });

    const nuevosProductos = filas
      .map(convertirProducto)
      .filter(producto => producto.producto !== "");

    if (nuevosProductos.length === 0) {
      throw new Error("No se encontraron productos en el catálogo.");
    }

    productos = nuevosProductos;
    ultimaCargaCorrecta = true;

    buscarProductos();
  } catch (error) {
    console.error("Error al cargar el catálogo:", error);

    if (!ultimaCargaCorrecta) {
      estado.textContent = `Error al cargar el catálogo: ${error.message}`;
    } else {
      console.warn("Se conservarán los datos de la última carga correcta.");
    }
  } finally {
    cargando = false;
  }
}

// Crea una línea de información con su etiqueta en negrita.
function agregarDato(tarjeta, etiqueta, valor) {
  if (!valor) return;

  const linea = document.createElement("p");
  const titulo = document.createElement("strong");

  titulo.textContent = `${etiqueta}: `;

  linea.append(titulo, document.createTextNode(valor));
  tarjeta.appendChild(linea);
}

// Formatea el precio en pesos colombianos.
function formatearPrecio(precio) {
  if (precio === null || precio < 0) {
    return "Precio por consultar";
  }

  return `$${precio.toLocaleString("es-CO", {
    maximumFractionDigits: 0
  })}`;
}

// Construye la tarjeta visual de cada producto.
function crearTarjeta(producto) {
  const tarjeta = document.createElement("article");
  tarjeta.className = "producto";

  // Encabezado: nombre a la izquierda y precio a la derecha.
  const encabezado = document.createElement("div");
  encabezado.className = "producto-encabezado";

  const nombre = document.createElement("h2");
  nombre.textContent = producto.producto;

  const precio = document.createElement("p");
  precio.className = "precio";
  precio.textContent = formatearPrecio(producto.precio);

  encabezado.append(nombre, precio);
  tarjeta.appendChild(encabezado);

  // Información adicional.
  agregarDato(tarjeta, "Categoría", producto.categoria);
  agregarDato(tarjeta, "Marca", producto.marca);

  // Disponibilidad según el inventario.
  const stock = document.createElement("span");

  if (producto.stock === null || producto.stock < 0) {
    stock.className = "agotado";
    stock.textContent = "Por confirmar";
  } else if (producto.stock > 0) {
    stock.className = "disponible";
    stock.textContent = "Disponible";
  } else {
    stock.className = "agotado";
    stock.textContent = "Agotado";
  }

  tarjeta.appendChild(stock);

  return tarjeta;
}

// Busca por nombre, marca o categoría y ordena por relevancia.
function buscarProductos() {
  resultados.replaceChildren();

  if (!productos.length) {
    if (ultimaCargaCorrecta) {
      estado.textContent = "No hay productos registrados.";
    }
    return;
  }

  const consulta = normalizar(campoBusqueda.value);

  if (!consulta) {
    estado.textContent =
      `Catálogo disponible: ${productos.length} productos.`;
    return;
  }

  const encontrados = productos
    .map(producto => {
      const nombre = normalizar(producto.producto);
      const marca = normalizar(producto.marca);
      const categoria = normalizar(producto.categoria);

      let prioridad = 99;

      if (nombre.startsWith(consulta)) {
        prioridad = 1;
      } else if (nombre.includes(consulta)) {
        prioridad = 2;
      } else if (marca.startsWith(consulta)) {
        prioridad = 3;
      } else if (categoria.startsWith(consulta)) {
        prioridad = 4;
      } else if (
        marca.includes(consulta) ||
        categoria.includes(consulta)
      ) {
        prioridad = 5;
      }

      return { producto, prioridad };
    })
    .filter(resultado => resultado.prioridad < 99)
    .sort((a, b) => a.prioridad - b.prioridad);

  estado.textContent = encontrados.length
    ? `Encontramos ${encontrados.length} resultado(s).`
    : "No encontramos ese producto.";

  // Mostrar máximo 3 productos por búsqueda
const productosLimitados = encontrados.slice(0, 3);

const tarjetas = productosLimitados.map(({ producto }) =>
  crearTarjeta(producto)
);

resultados.append(...tarjetas);
}

// Eventos del buscador.
formulario.addEventListener("submit", evento => {
  evento.preventDefault();
  buscarProductos();
});

campoBusqueda.addEventListener("input", buscarProductos);

// Carga inicial y actualización periódica.
cargarCatalogo();
setInterval(cargarCatalogo, 60_000);
