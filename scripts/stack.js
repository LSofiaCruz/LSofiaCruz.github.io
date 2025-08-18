import { stack } from "./models.js";

let stack_dictionary = {}

function cargarStack(stack) {
    let contenedor = document.getElementById("stack-contenedor");

    // stack.forEach(item => {processStack(item)})

    const stackHard = stack.filter(skill => skill.grupo === "hard")
    const stackSoft = stack.filter(skill => skill.grupo === "soft")
    const stackIdiom = stack.filter(skill => skill.grupo === "idiomas")

    contenedor.appendChild(createStackSection("HARD", stackHard));
    contenedor.appendChild(createStackSection("SOFT", stackSoft));
    contenedor.appendChild(createStackSection("IDIOMAS", stackIdiom));
}

// function processStack(stack) {
//     if(stack_dictionary[stack.grupo] === undefined) {
//         const stack_key = stack.grupo
//         const stack_value = stack.nombre
//         const key_value = {
//             `${stack_key}`: [stack_value]
//         }
//         stack_dictionary.add({
//              key_value
//         })
//     } else {
//         console.log("DICT::::", stack_dictionary[stack.grupo])
//     }
// }

function renderStack(stackGroup, container) {
    stackGroup.forEach(tech => {
        const item = document.createElement("div");
        item.classList.add("stack-item");

        item.innerHTML = `
        <div class="circular-progress" style="--percentage: ${tech.nivel}; --color: ${tech.color};">
          ${tech.nivel}%
        </div>
        <p class="work-sans-bold"><strong>${tech.nombre}</strong></p>
        <p class="work-sans-ligth">${tech.nivel_txt}</p>
      `;

        container.appendChild(item);
    });
    return container;
}

function createStackSection(tittle, stack){
    let group = document.createElement('div');
    group.classList.add("stack-group");

    group = renderStack(stack, group);

    const section = document.createElement('div');

    section.classList.add("stack-tittle");
    section.innerHTML = `<h1  class="work-sans-bold vertical-text">${tittle}</h1>`;
    section.appendChild(group);

    return section;
}

document.addEventListener("DOMContentLoaded", () => {
    cargarStack(stack);
});
