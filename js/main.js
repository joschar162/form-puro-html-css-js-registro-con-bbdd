// Variables globales
let registros = [];
let editandoId = null;
let registroAEditar = null;

// Elementos del DOM
const form = document.getElementById("registration-form");
const alertError = document.getElementById("alert-error");
const alertSuccess = document.getElementById("alert-success");
const alertInfo = document.getElementById("alert-info");
const alertWarning = document.getElementById("alert-warning");
const alertDuplicate = document.getElementById("alert-duplicate");
const submitBtn = document.getElementById("submit-btn");
const tableBody = document.getElementById("registrados-body");
const emptyState = document.getElementById("empty-state");
const table = document.getElementById("registrados-table");
const totalRegistrados = document.getElementById("total-registrados");
const modalVerificacion = document.getElementById("modal-verificacion");
const modalError = document.getElementById("modal-error");
const codeInputs = document.querySelectorAll(".code-input");

// Cargar registros desde localStorage al iniciar
function inicializar() {
  // Escuchar cambios en tiempo real desde Firebase
  db.ref("inscritos").on("value", (snapshot) => {
    const data = snapshot.val();

    // Convertir el objeto de Firebase a un array para tu tabla
    if (data) {
      registros = Object.keys(data).map((key) => ({
        id: key,
        ...data[key],
      }));
    } else {
      registros = [];
    }

    renderizarTabla();
    console.log("Datos sincronizados con Firebase");
  });
}

// Generar ID único
function generarId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// Busca el botón por ID y asigna el evento
document.getElementById("btnExportar").addEventListener("click", () => {
  // Verificación de seguridad
  if (typeof XLSX === "undefined") {
    alert(
      "La librería de Excel aún no ha cargado. Por favor, refresca la página."
    );
    return;
  }

  try {
    const tablaOriginal = document.getElementById("registrados-table");

    if (registros.length === 0) {
      alert("No hay residentes registrados para exportar.");
      return;
    }

    // Crear copia limpia sin la columna de "Acciones"
    const tablaClonada = tablaOriginal.cloneNode(true);
    const filas = tablaClonada.querySelectorAll("tr");
    filas.forEach((fila) => {
      if (fila.lastElementChild) {
        fila.removeChild(fila.lastElementChild);
      }
    });

    // Generar y descargar el archivo
    const wb = XLSX.utils.table_to_book(tablaClonada, { sheet: "Inscritos" });
    XLSX.writeFile(wb, "Lista_padron_socios.xlsx");
  } catch (error) {
    console.error("Error al exportar:", error);
  }
});
// Normalizar texto (quitar acentos y espacios extras)
function normalizarTexto(texto) {
  return texto
    .toString()
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

// Verificar si ya existe el registro
function verificarDuplicado(
  nombre,
  apellidoPaterno,
  apellidoMaterno,
  excludeId = null
) {
  const nombreNormalizado = normalizarTexto(nombre);
  const paternoNormalizado = normalizarTexto(apellidoPaterno);
  const maternoNormalizado = normalizarTexto(apellidoMaterno);

  return registros.find((registro) => {
    if (excludeId && registro.id === excludeId) return false;

    const nomMatch = normalizarTexto(registro.nombre) === nombreNormalizado;
    const patMatch =
      normalizarTexto(registro.apellidoPaterno) === paternoNormalizado;
    const matMatch =
      normalizarTexto(registro.apellidoMaterno) === maternoNormalizado;

    return nomMatch && patMatch && matMatch;
  });
}

// Mostrar alerta
function mostrarAlerta(tipo, mensaje) {
  // Ocultar todas las alertas primero
  alertError.style.display = "none";
  alertSuccess.style.display = "none";
  alertInfo.style.display = "none";
  alertWarning.style.display = "none";
  alertDuplicate.style.display = "none";

  const alerta = document.getElementById(`alert-${tipo}`);
  if (alerta) {
    alerta.textContent = mensaje;
    alerta.style.display = "block";

    // Auto-ocultar después de 5 segundos (excepto error)
    if (tipo !== "error" && tipo !== "duplicate" && tipo !== "warning") {
      setTimeout(() => {
        alerta.style.display = "none";
      }, 5000);
    }
  }
}

// Obtener nombre completo formateado
function obtenerNombreCompleto(registro) {
  return `${registro.nombre} ${registro.apellidoPaterno} ${registro.apellidoMaterno}`;
}

// Obtener badge de género
function obtenerBadgeGenero(genero) {
  const clases =
    genero === "masculino" ? "gender-masculino" : "gender-femenino";
  const texto = genero === "masculino" ? "Masculino" : "Femenino";
  return `<span class="gender-badge ${clases}">${texto}</span>`;
}

// Renderizar tabla de registros
function renderizarTabla() {
  // Actualizar contador
  totalRegistrados.textContent = registros.length;

  // Mostrar/ocultar estado vacío
  if (registros.length === 0) {
    table.style.display = "none";
    emptyState.style.display = "block";
  } else {
    table.style.display = "table";
    emptyState.style.display = "none";

    // Ordenar alfabéticamente
    const registrosOrdenados = [...registros].sort((a, b) => {
      const nombreA = obtenerNombreCompleto(a);
      const nombreB = obtenerNombreCompleto(b);
      return nombreA.localeCompare(nombreB);
    });

    tableBody.innerHTML = registrosOrdenados
      .map(
        (registro) => `
                    <tr>
                        <td><strong>${obtenerNombreCompleto(
                          registro
                        )}</strong></td>
                        <td>${registro.celular}</td>
                        <td>${obtenerBadgeGenero(registro.genero)}</td>
                        <td>
                            <button class="btn-edit" onclick="solicitarVerificacion('${
                              registro.id
                            }')">
                                ✏️ Editar
                            </button>
                        </td>
                    </tr>
                `
      )
      .join("");
  }

  // Guardar en localStorage
  // localStorage.setItem("registrosHuaylia", JSON.stringify(registros));
}

// Solicitar verificación antes de editar
window.solicitarVerificacion = function (id) {
  registroAEditar = registros.find((r) => r.id === id);
  if (!registroAEditar) return;

  // Limpiar campos del código
  codeInputs.forEach((input) => {
    input.value = "";
    input.classList.remove("error");
  });
  modalError.classList.remove("error");
  modalError.textContent = "";

  // Enfocar el primer campo
  codeInputs[0].focus();

  // Mostrar modal
  modalVerificacion.classList.add("active");
};

// Cancelar verificación
function cancelarVerificacion() {
  modalVerificacion.classList.remove("active");
  registroAEditar = null;
}

// Confirmar verificación
function confirmarVerificacion() {
  // Obtener código ingresado
  let codigoIngresado = "";
  codeInputs.forEach((input) => {
    codigoIngresado += input.value;
  });

  // Validar que se ingresaron los 4 dígitos
  if (codigoIngresado.length !== 4) {
    modalError.textContent = "⚠️ Por favor, ingrese los 4 dígitos del código.";
    modalError.classList.add("error");
    codeInputs.forEach((input) => {
      if (!input.value) input.classList.add("error");
    });
    return;
  }

  // Verificar código
  if (registroAEditar && registroAEditar.codigoAcceso === codigoIngresado) {
    // Código correcto, proceder con edición
    modalVerificacion.classList.remove("active");
    editarRegistro(registroAEditar.id);
    registroAEditar = null;
  } else {
    // Código incorrecto
    modalError.textContent =
      "❌ Código de acceso incorrecto. Verifique e intente nuevamente.";
    modalError.classList.add("error");

    // Vibrar en dispositivos móviles
    if (navigator.vibrate) {
      navigator.vibrate(200);
    }

    // Limpiar campos y reenfocar
    codeInputs.forEach((input) => {
      input.value = "";
      input.classList.add("error");
    });
    codeInputs[0].focus();
  }
}

// Editar registro (después de verificación)
function editarRegistro(id) {
  const registro = registros.find((r) => r.id === id);
  if (!registro) return;

  // Llenar formulario con datos existentes
  document.getElementById("nombre").value = registro.nombre;
  document.getElementById("apellido_paterno").value = registro.apellidoPaterno;
  document.getElementById("apellido_materno").value = registro.apellidoMaterno;
  document.getElementById("celular").value = registro.celular;
  document.getElementById("genero").value = registro.genero;
  document.getElementById("codigo_acceso").value = registro.codigoAcceso;

  // Configurar modo edición
  editandoId = id;
  submitBtn.textContent = "✏️ Actualizar Registro";
  submitBtn.classList.remove("btn-primary");
  submitBtn.classList.add("btn-success");

  // Scroll al formulario
  form.scrollIntoView({ behavior: "smooth", block: "start" });

  // Quitar foco del último campo
  document.getElementById("codigo_acceso").blur();

  // Mostrar mensaje informativo
  mostrarAlerta(
    "info",
    '✏️ Modo edición: Modifique los datos y haga clic en "Actualizar Registro"'
  );
}

// Resetear formulario
function resetearFormulario() {
  form.reset();
  editandoId = null;
  submitBtn.textContent = "Registrarse";
  submitBtn.classList.remove("btn-success");
  submitBtn.classList.add("btn-primary");
}

// Manejar envío del formulario
form.addEventListener("submit", function (e) {
  e.preventDefault();

  // Obtener valores
  const nombre = document.getElementById("nombre").value.trim();
  const apellidoPaterno = document
    .getElementById("apellido_paterno")
    .value.trim();
  const apellidoMaterno = document
    .getElementById("apellido_materno")
    .value.trim();
  const celular = document.getElementById("celular").value.trim();
  const genero = document.getElementById("genero").value;
  const codigoAcceso = document.getElementById("codigo_acceso").value;

  // Validaciones
  if (
    !nombre ||
    !apellidoPaterno ||
    !apellidoMaterno ||
    !celular ||
    !genero ||
    !codigoAcceso
  ) {
    mostrarAlerta(
      "error",
      "Por favor, complete todos los campos obligatorios."
    );
    return;
  }

  // Validar formato de celular (9 dígitos para Perú)
  if (!/^9\d{8}$/.test(celular)) {
    mostrarAlerta(
      "error",
      "El número de celular debe tener exactamente 9 dígitos y empezar con 9."
    );
    return;
  }

  // Validar código de acceso (4 dígitos)
  if (!/^\d{4}$/.test(codigoAcceso)) {
    mostrarAlerta(
      "error",
      "El código de acceso debe tener exactamente 4 dígitos numéricos."
    );
    return;
  }

  // Verificar duplicado (solo si no estamos editando el mismo registro)
  const duplicado = verificarDuplicado(
    nombre,
    apellidoPaterno,
    apellidoMaterno,
    editandoId
  );
  if (duplicado) {
    const nombreCompleto = `${nombre} ${apellidoPaterno} ${apellidoMaterno}`;
    mostrarAlerta(
      "error",
      `⚠️ "${nombreCompleto}" ya está registrado en la lista.`
    );
    return;
  }

  if (editandoId) {
    // ACTUALIZAR EN FIREBASE
    db.ref("inscritos/" + editandoId)
      .update({
        nombre,
        apellidoPaterno,
        apellidoMaterno,
        celular,
        genero,
        codigoAcceso,
      })
      .then(() => {
        mostrarAlerta("success", "✅ Registro actualizado en la nube.");
      });
  } else {
    // CREAR NUEVO EN FIREBASE
    const nuevoId = generarId();
    const nuevoRegistro = {
      nombre,
      apellidoPaterno,
      apellidoMaterno,
      celular,
      genero,
      codigoAcceso,
    };

    db.ref("inscritos/" + nuevoId)
      .set(nuevoRegistro)
      .then(() => {
        mostrarAlerta("success", "✅ ¡Registro exitoso en la nube!");
      });
  }

  // Renderizar y resetear
  renderizarTabla();
  resetearFormulario();
});

// Validación en tiempo real para campos numéricos
document.getElementById("celular").addEventListener("input", function (e) {
  this.value = this.value.replace(/\D/g, "").slice(0, 9);
});

document
  .getElementById("codigo_acceso")
  .addEventListener("input", function (e) {
    this.value = this.value.replace(/\D/g, "").slice(0, 4);
  });

// Navegación entre campos del código de verificación
codeInputs.forEach((input, index) => {
  input.addEventListener("input", function (e) {
    // Remover error
    this.classList.remove("error");
    modalError.classList.remove("error");
    modalError.textContent = "";

    // Mover al siguiente campo si hay valor
    if (this.value && index < codeInputs.length - 1) {
      codeInputs[index + 1].focus();
    }
  });

  input.addEventListener("keydown", function (e) {
    // Retroceder al campo anterior con backspace
    if (e.key === "Backspace" && !this.value && index > 0) {
      codeInputs[index - 1].focus();
    }

    // Permitir navegación con flechas
    if (e.key === "ArrowLeft" && index > 0) {
      codeInputs[index - 1].focus();
    }
    if (e.key === "ArrowRight" && index < codeInputs.length - 1) {
      codeInputs[index + 1].focus();
    }
  });

  // Seleccionar todo al hacer clic
  input.addEventListener("click", function (e) {
    this.select();
  });
});

// Toggle password visibility
function togglePassword(inputId, button) {
  const input = document.getElementById(inputId);
  if (input.type === "password") {
    input.type = "text";
    button.textContent = "🔒";
  } else {
    input.type = "password";
    button.textContent = "👁️";
  }
}

// Cerrar modal con Escape
document.addEventListener("keydown", function (e) {
  if (e.key === "Escape" && modalVerificacion.classList.contains("active")) {
    cancelarVerificacion();
  }
});

// Cerrar modal al hacer clic fuera
modalVerificacion.addEventListener("click", function (e) {
  if (e.target === modalVerificacion) {
    cancelarVerificacion();
  }
});

// Inicializar al cargar la página
document.addEventListener("DOMContentLoaded", inicializar);
