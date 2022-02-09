console.log("Hello world!");

//const circlePerimeter = (r) => {return 2*3.14*r}
//const circlePerimeter = r => {return 2*3.14*r}
const circlePerimeter = r => 2 * 3.14 * r

console.log("circlePerimeter(4)", circlePerimeter(4));

const array1 = ['Apple', 'Orange', 'Banana'];

//clone => faz copia para outra posição para memoria
const cloneArray = [...array1]
cloneArray.push("Peaches");

console.log("cloneArray:", cloneArray);

const marco = {
    "name": "Marco",
    "tech": ["Python", "React", "NodeJs"]
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

const p = new Promise((resolve, reject) => {
    setTimeout(() => {
        const result = circlePerimeter(2)
        resolve(result);
    }, 3000);
});

p.then((result) => { console.log("result = ", result) });

const circleArraySync = (r) => new Promise((resolve, reject) => {
    setTimeout(() => { resolve(circlePerimeter(r)) }, 3000);
})

const calculateArea = async () => {
    try {
        const result = await circleArraySync(10);
        console.log("result = ", result);
    }
    catch (e) {
        console.log("error=", e);
    }
}

const calculateAreaMultipleCalls = async () => {
    try {
        const promises = [circleArraySync(1), circleArraySync(2), circleArraySync(3), circleArraySync(4)]

        const results = await Promise.all(promises)
        
        console.log("results = ", results)
    }
    catch (e) {
        console.log("error=", e);
    }
}

calculateArea();
calculateAreaMultipleCalls();