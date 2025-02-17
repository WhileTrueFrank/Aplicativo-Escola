document.addEventListener("DOMContentLoaded", function () {
    document.getElementById("reservaForm").addEventListener("submit", async function (event) {
        event.preventDefault();

        const nome = document.getElementById("nome").value;
        const quantidade = document.getElementById("quantidade").value;
        const data = document.getElementById("data").value;
        const hora_inicio = document.getElementById("hora_inicio").value;
        const hora_fim = document.getElementById("hora_fim").value;

        const response = await fetch("/reservar", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                nome,
                quantidade,
                data,
                hora_inicio,
                hora_fim
            })
        });

        const result = await response.json();
        if (response.ok) {
            alert("Reserva realizada com sucesso!");
            document.getElementById("reservaForm").reset();
        } else {
            alert("Erro: " + result.error); // Exibe erro caso já exista reserva
        }
    });
});
