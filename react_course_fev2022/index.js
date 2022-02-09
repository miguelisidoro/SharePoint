console.log("Hello world!");

//const circlePerimeter = (r) => {return 2*3.14*r}
//const circlePerimeter = r => {return 2*3.14*r}
const circlePerimeter = r => 2*3.14*r

console.log ("circlePerimeter(4)", circlePerimeter(4));

const array1 = ['Apple', 'Orange', 'Banana'];

//clone => faz copia para outra posição para memoria
const cloneArray = [...array1]
cloneArray.push("Peaches");

console.log ("cloneArray:", cloneArray);

const marco = {
    "name": "Marco",
    "tech": ["Python","React","NodeJs"]
}

//sread operator -> n faz clone, usa a mesma operação de memoria
const pedro = {
    ...marco,
    name: "Pedro"
}

console.log("marco = ", marco);
console.log("pedro = ", pedro);

pedro.tech = [...marco.tech, "Flutter"];

console.log("pedro = ", pedro);