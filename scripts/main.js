import { jobs } from "./models.js"

document.addEventListener("scroll", () => {
  const sections = document.querySelectorAll("section");
  const scrollY = window.scrollY;
  sections.forEach(sec => {
    const top = sec.offsetTop - 100;
    const height = sec.offsetHeight;
    const id = sec.getAttribute("id");
    const link = document.querySelector(`nav a[href="#${id}"]`);
    if (scrollY >= top && scrollY < top + height) {
      link.classList.add("active");
    } else {
      link.classList.remove("active");
    }
  });
});

function cargarExperiencia(data) {
  const contenedor = document.getElementById("experiencia-contenedor");

  data.carrera.forEach(item => {
    const div = document.createElement("div");
    div.classList.add("job");
    div.classList.add(`border-left-${item.color}`)

    div.innerHTML = `
        <div class="card-tittle work-sans-bold">
            <h3>${item.empresa}</h3>
            <span class="skills work-sans-regular">${item.habilidades ? item.habilidades : ''}</span>
        </div>
        <p class="work-sans-bold"><strong>${item.cargo}</strong> (${item.fecha_inicio} - ${item.fecha_fin})</p>
        <p class="work-sans-ligth">${item.descripcion}</p>
      `;

    contenedor.appendChild(div);
  });
}

document.addEventListener("DOMContentLoaded", () => {
  cargarExperiencia(jobs);
});
