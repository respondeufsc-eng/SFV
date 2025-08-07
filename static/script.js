let currentStep = '';

function nextStep(step) {
    currentStep = step;
    const resultDiv = document.getElementById('result');
    switch (step) {
        case 'cleaning':
            resultDiv.innerHTML = `<p>O módulos foram limpos?</p>
            <button onclick="nextStep('visualInspection')">Sim</button>
            <button onclick="nextStep('goClean')">Não</button>`;
            break;
        case 'goClean':
            resultDiv.innerHTML = `
            <p>Antes de iniciar, favor limpar os módulos.</p>
            <p>Depois clique para continuar</p>
            <button onclick="nextStep('cleaning')">Continuar</button>`;
            break;
        case 'visualInspection':
            resultDiv.innerHTML = `<p>A inspeção visual (Visual Inspection) mostra vidro quebrado?</p>
            <button onclick="recycle()">Sim</button>
            <button onclick="handleGlassDamage('no')">Não</button>`;
            break;

        case 'maintenance':
            resultDiv.innerHTML = `<p>Qual é a gravidade dos danos?</p>
            <button onclick="nextStep('severeDamage')">Graves</button>
            <button onclick="nextStep('fixable')">Pode ser consertado</button>`;
            break;

        case 'severeDamage':
            recycle();
            break;

        case 'fixable':
            nextStep('insulationResistanceTest');
            break;

        case 'insulationResistanceTest':
            resultDiv.innerHTML = `<p>O teste de resistência de isolamento deu mais que 40 MΩ·m²?</p>
            <button onclick="nextStep('ivCurveTest')">Sim</button>
            <button onclick="recycle()">Não</button>`;
            break;

        case 'ivCurveTest': //Colocar o que deve ser importante neste teste, quais valores devem ser levados em consideração
            resultDiv.innerHTML = `<p>Você sabe a idade do módulo?</p>
            <button onclick="nextStep('moduleAgeKnown')">Sim</button>
            <button onclick="nextStep('moduleAgeUnknown')">Não</button>`;
            break;

        case 'moduleAgeKnown':
            resultDiv.innerHTML = `<p>O módulo entrega até 10% da potência esperada?</p>
            <button onclick="nextStep('expectedPowerCheck')">Sim</button>
            <button onclick="recycle()">Não</button>`;
            break;

        case 'moduleAgeUnknown':
            resultDiv.innerHTML = `<p>O módulo entrega mais de 60% da potência original?</p>
            <button onclick="nextStep('expectedPowerCheck')">Sim</button>
            <button onclick="recycle()">Não</button>`;
            break;

        case 'expectedPowerCheck':
            resultDiv.innerHTML = `<p>O módulo entrega a potência esperada?</p>
            <button onclick="nextStep('classA')">Sim</button>
            <button onclick="nextStep('classB')">Não</button>`;
            break;

        case 'classA': // Informar a diferença entre clasess
        case 'classB':
            nextStep('electroluminescenceTest');
            break;

        case 'electroluminescenceTest':
            resultDiv.innerHTML = `<p>O teste de eletroluminescência encontrou rachaduras?</p>
            <button onclick="recycle()">Sim</button>
            <button onclick="nextStep('checkDamagedCells')">Não</button>`;
            break;

        case 'checkDamagedCells':
            resultDiv.innerHTML = `<p>Mais de 50% das células estão danificadas?</p>
            <button onclick="recycle()">Sim</button>
            <button onclick="secondLife()">Não</button>`;
            break;

        default:
            resultDiv.innerHTML = `<p>Etapa não reconhecida.</p>`;
    }
}

function handleGlassDamage(answer) {
    const resultDiv = document.getElementById('result');
    if (answer === 'no') {
        resultDiv.innerHTML = `<p>Verificando outros componentes: backsheet, caixa de junção, cabos e conectores estão danificados?</p>
        <button onclick="nextStep('fixable')">Pode ser consertado</button>
        <button onclick="nextStep('insulationResistanceTest')">Não tem danos</button>
        <button onclick="nextStep('severeDamage')">Não pode ser consertado</button>`;
    } else {
        resultDiv.innerHTML = `<p>Verificando outros componentes: backsheet, caixa de junção, cabos e conectores...</p>`;
        setTimeout(() => nextStep('maintenance'), 1000);
    }
}

function recycle() {
    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = `<p>Resultado: <strong>Reciclagem (R)</strong> ♻️</p>`;
}

function secondLife() {
    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = `<p>Resultado: <strong>Disponível para Segunda Vida (SL)</strong> 🌱</p>`;
}

function generateExcel() {
    const quantity = document.getElementById('quantity').value;
    const url = `/generate_excel?quantity=${quantity}`;
    window.location.href = url;
}

