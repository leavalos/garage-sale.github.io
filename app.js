document.addEventListener('DOMContentLoaded', function () {
    // Cambia 'nombre-del-archivo.xlsx' por el nombre de tu archivo Excel
    var excelFilePath = 'Garage sale.xlsx';
    // Llamada a la función para cargar y procesar el archivo Excel
    loadExcel(excelFilePath);
});

function loadExcel(filePath) {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', filePath, true);
    xhr.responseType = 'arraybuffer';

    xhr.onload = function (e) {
        var arraybuffer = xhr.response;

        /* Convertir datos binarios a una hoja de cálculo de JavaScript */
        var data = new Uint8Array(arraybuffer);
        var arr = new Array();
        for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
        var bstr = arr.join("");
        var workbook = XLSX.read(bstr, { type: 'binary' });

        /* Obtener la primera hoja de cálculo */
        var firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        /* Convertir la hoja de cálculo a un arreglo de objetos */
        var products = XLSX.utils.sheet_to_json(firstSheet);

        /* Mostrar productos en la página */
        displayProducts(products);
    };

    xhr.send();
}

function displayProducts(products) {
    var productListContainer = document.getElementById('product-list');

    var productContainer;
    
    products.forEach(function (product, index) {
        // Crea un nuevo contenedor cada cinco elementos
        if (index % 5 === 0) {
            productContainer = document.createElement('div');
            productContainer.classList.add('product-container');
            productListContainer.appendChild(productContainer);
        }

        var productDiv = document.createElement('div');
        productDiv.classList.add('product');

        var productImage = document.createElement('img');
        productImage.src = "https://drive.google.com/thumbnail?id=" + getIdImage(product['Imagen del producto']); // Asegúrate de que la columna se llame 'Imagen'
        productImage.alt = product['Nombre'];

        var productName = document.createElement('h3');
        productName.textContent = product['Nombre'];

        var productDetails = document.createElement('div');
        productDetails.classList.add('details');

        var productBrand = document.createElement('p');
        productBrand.textContent = `Marca: ${product['Marca']}`;

        var productCategory = document.createElement('p');
        productCategory.textContent = `Categoría: ${product['Categoria']}`;

        var productSize = document.createElement('p');
        productSize.textContent = `Talle: ${product['Talle']}`;

        var productPrice = document.createElement('p');
        productPrice.textContent = `Precio: ${product['Precio ']}`;

        productDetails.appendChild(productBrand);
        productDetails.appendChild(productCategory);
        productDetails.appendChild(productSize);

        productDiv.appendChild(productImage);
        productDiv.appendChild(productName);
        productDiv.appendChild(productDetails);
        productDiv.appendChild(productPrice);

        productContainer.appendChild(productDiv);
    });
}

function getIdImage(link) {

    var regex = /\/file\/d\/([^\/]+)\//;
    var match = link.match(regex);

    if (match) {
        return match[1]
    }else {
        return ""
    }
    
}