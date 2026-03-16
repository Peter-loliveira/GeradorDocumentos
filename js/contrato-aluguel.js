// contrato-aluguel.js - Gerador de Contrato de Locação
// Versão modificada: Usa MODELO.docx como template, fonte Calibri, e nome específico

function coletarDadosAluguel() {
    return {
        locador: {
            nome: document.getElementById('locador_nome').value.toUpperCase(),
            nacionalidade: document.getElementById('locador_nacionalidade').value.toLowerCase(),
            estado_civil: document.getElementById('locador_estado_civil').value,
            rg: document.getElementById('locador_rg').value,
            rg_orgao: document.getElementById('locador_rg_orgao').value,
            cpf: document.getElementById('locador_cpf').value,
            endereco: document.getElementById('locador_endereco').value,
            email: document.getElementById('locador_email').value,
            telefone: document.getElementById('locador_telefone').value
        },
        locatario: {
            nome: document.getElementById('locatario_nome').value.toUpperCase(),
            nacionalidade: document.getElementById('locatario_nacionalidade').value.toLowerCase(),
            estado_civil: document.getElementById('locatario_estado_civil').value,
            rg: document.getElementById('locatario_rg').value,
            rg_orgao: document.getElementById('locatario_rg_orgao').value,
            cpf: document.getElementById('locatario_cpf').value,
            endereco: document.getElementById('locatario_endereco').value,
            email: document.getElementById('locatario_email').value,
            telefone: document.getElementById('locatario_telefone').value
        },
        corretor: {
            nome: document.getElementById('corretor_nome').value.toUpperCase(),
            cpf: document.getElementById('corretor_cpf').value,
            creci: document.getElementById('corretor_creci').value
        },
        contrato: {
            data_pagamento: document.getElementById('data_pagamento').value,
            data_entrada: document.getElementById('data_entrada').value,
            tipo_seguranca: document.querySelector('input[name="tipo_seguranca"]:checked')?.value,
            caucao_valor: document.getElementById('caucao_valor')?.value || '',
            caucao_meses: document.getElementById('caucao_meses')?.value || '',
            tipo_imovel: document.getElementById('tipo_imovel').value,
            imovel_descricao: document.getElementById('imovel_descricao').value,
            imovel_endereco: document.getElementById('imovel_endereco').value,
            valor_aluguel: document.getElementById('valor_aluguel').value,
            prazo_locacao: document.getElementById('prazo_locacao').value
        }
    };
}

function validarCamposAluguel() {
    const campos = [
        'locador_nome', 'locador_nacionalidade', 'locador_estado_civil', 'locador_rg',
        'locador_rg_orgao', 'locador_cpf', 'locador_endereco', 'locador_email', 'locador_telefone',
        'locatario_nome', 'locatario_nacionalidade', 'locatario_estado_civil', 'locatario_rg',
        'locatario_rg_orgao', 'locatario_cpf', 'locatario_endereco', 'locatario_email', 'locatario_telefone',
        'corretor_nome', 'corretor_cpf', 'corretor_creci',
        'data_pagamento', 'data_entrada',
        'tipo_imovel', 'imovel_descricao', 'imovel_endereco',
        'valor_aluguel', 'prazo_locacao'
    ];

    for (let campo of campos) {
        const el = document.getElementById(campo);
        if (!el || !el.value.trim()) {
            if (el) {
                el.focus();
                el.style.borderColor = '#E31837';
                setTimeout(() => el.style.borderColor = '', 3000);
            }
            return false;
        }
    }

    const tipoSeg = document.querySelector('input[name="tipo_seguranca"]:checked');
    if (!tipoSeg) {
        alert('Por favor, selecione o tipo de segurança do imóvel.');
        return false;
    }

    if (tipoSeg.value === 'caucao') {
        if (!document.getElementById('caucao_valor').value || !document.getElementById('caucao_meses').value) {
            alert('Por favor, preencha os dados da caução.');
            return false;
        }
    }

    return true;
}

// Função para escapar XML
function escapeXml(text) {
    if (!text) return '';
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

// Função para criar um parágrafo Word XML com fonte Calibri
function createWordParagraph(text, options = {}) {
    const { bold = false, alignment = 'both', spacing = '200', size = '24' } = options;

    let alignMap = {
        'left': 'start',
        'right': 'end',
        'center': 'center',
        'both': 'both'
    };

    let jc = alignMap[alignment] || 'both';

    // Criar runs (partes do texto)
    let runsXml = '';

    if (Array.isArray(text)) {
        // Array de objetos {text, bold}
        runsXml = text.map(t => {
            const b = t.bold || bold ? '<w:b/>' : '';
            return `<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="${t.size || size}"/><w:szCs w:val="${t.size || size}"/>${b}</w:rPr><w:t xml:space="preserve">${escapeXml(t.text)}</w:t></w:r>`;
        }).join('');
    } else {
        const b = bold ? '<w:b/>' : '';
        runsXml = `<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/>${b}</w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
    }

    return `<w:p><w:pPr><w:jc w:val="${jc}"/><w:spacing w:after="${spacing}"/></w:pPr>${runsXml}</w:p>`;
}

// Função para criar título de cláusula
function createClauseTitle(title) {
    return createWordParagraph(title, { bold: true, size: '24', spacing: '200' });
}

// Função para criar texto de cláusula com partes bold
function createClauseText(parts) {
    const textArray = parts.map(p => ({
        text: p.text,
        bold: p.bold || false,
        size: '24'
    }));
    return createWordParagraph(textArray, { alignment: 'both', spacing: '200' });
}

// Função principal para gerar o XML do contrato
function generateContractXml(dados) {
    const dataEntrada = new Date(dados.contrato.data_entrada);
    const dataFim = new Date(dataEntrada);
    dataFim.setMonth(dataFim.getMonth() + parseInt(dados.contrato.prazo_locacao));

    const mesesExtenso = numeroPorExtenso(parseInt(dados.contrato.prazo_locacao));
    const caucaoMesesExtenso = dados.contrato.caucao_meses ? numeroPorExtenso(parseInt(dados.contrato.caucao_meses)) : '';

    let xml = '';

    // Título
    xml += createWordParagraph('CONTRATO DE LOCAÇÃO DE IMÓVEL RESIDENCIAL', {
        bold: true,
        alignment: 'center',
        size: '28',
        spacing: '400'
    });

    // LOCADOR
    xml += createClauseTitle('LOCADOR(A)');
    xml += createClauseText([
        { text: dados.locador.nome + ', ', bold: true },
        { text: dados.locador.nacionalidade + ', ' + dados.locador.estado_civil.toLowerCase() + ', portador(a) do ' },
        { text: 'RG nº ' + dados.locador.rg + ', ', bold: true },
        { text: 'emitido pela ' + dados.locador.rg_orgao + ', inscrito(a) no ' },
        { text: 'CPF sob nº ' + dados.locador.cpf + ' ', bold: true },
        { text: 'residente e domiciliado(a) na ' + dados.locador.endereco + ', Telefone: ' + dados.locador.telefone + ', E-mail: ' + dados.locador.email + ', representado(a) neste ato pelo(a) corretor(a) ' },
        { text: dados.corretor.nome + ', ', bold: true },
        { text: 'corretor(a) de imóveis, inscrito(a) no ' },
        { text: 'CPF sob o nº ' + dados.corretor.cpf + ', ', bold: true },
        { text: 'registrado(a) no ' },
        { text: 'CRECI-BA sob o nº ' + dados.corretor.creci + ', ', bold: true },
        { text: 'telefone de contato +55 (71) 999441701, e-mail peteroliveira@remax.com.br.' }
    ]);

    // LOCATÁRIO
    xml += createClauseTitle('LOCATÁRIO(A)');
    xml += createClauseText([
        { text: dados.locatario.nome + ', ', bold: true },
        { text: dados.locatario.nacionalidade + ', ' + dados.locatario.estado_civil.toLowerCase() + ', portador(a) do ' },
        { text: 'RG nº ' + dados.locatario.rg + ', ', bold: true },
        { text: 'emitido pela ' + dados.locatario.rg_orgao + ', inscrito(a) no ' },
        { text: 'CPF sob nº ' + dados.locatario.cpf + ', ', bold: true },
        { text: 'residente e domiciliado(a) na ' + dados.locatario.endereco + ', telefone ' + dados.locatario.telefone + ', e-mail ' + dados.locatario.email + '.' }
    ]);

    // IMOBILIÁRIA
    xml += createClauseTitle('IMOBILIÁRIA');
    xml += createWordParagraph(
        'JAUÁ IMÓVEIS E EMPREENDIMENTOS LTDA. (RE/MAX Litorânea), inscrita no CNPJ sob o nº 07.788.314.0001-40, registrada no CRECI-BA sob o nº 1101 PJ, com sede na Rua Direta de Jaua, Loja 4, Jaua, Camaçari, Bahia, CEP: 42828-576, com telefone de contato +55 (71) 3672-1664 e e-mail litoranea@remax.com.br.',
        { alignment: 'both', spacing: '300' }
    );

    // CORRETOR
    xml += createClauseTitle('CORRETOR');
    xml += createClauseText([
        { text: dados.corretor.nome + ', ', bold: true },
        { text: 'corretor de imóveis, inscrito no CPF sob o nº ' + dados.corretor.cpf + ', registrada no CRECI-BA sob o nº ' + dados.corretor.creci + ', telefone de contato +55 (71) 9-9944-1701, e-mail peteroliveira@remax.com.br.' }
    ]);

    // Introdução
    xml += createClauseText([
        { text: 'As partes acima qualificadas estabelecem entre si o presente ' },
        { text: 'CONTRATO DE LOCAÇÃO DE IMÓVEL RESIDENCIAL ', bold: true },
        { text: 'mediante as condições e cláusulas seguintes:' }
    ]);

    // CLÁUSULA 1
    xml += createClauseTitle('CLÁUSULA 1ª – OBJETO DO CONTRATO');
    xml += createClauseText([
        { text: '1. O presente instrumento tem por ' },
        { text: 'OBJETO ', bold: true },
        { text: 'o imóvel do tipo ' },
        { text: dados.contrato.tipo_imovel.toLowerCase() + ' ', bold: true },
        { text: 'de propriedade do ' },
        { text: 'LOCADOR(A)', bold: true },
        { text: ', situado na ' + dados.contrato.imovel_endereco + ', ' + dados.contrato.imovel_descricao + '.' }
    ]);

    xml += createClauseText([
        { text: 'Parágrafo primeiro: ', bold: true },
        { text: 'Quando do início da locação será lavrado laudo de vistoria no qual constará, pormenorizadamente, a descrição da quantidade, qualidade e espécies de móveis e utensílios existentes, bem como do estado de conservação do imóvel, suas instalações hidráulicas e elétricas.' }
    ]);

    xml += createClauseText([
        { text: 'Parágrafo segundo: ', bold: true },
        { text: 'A presente ' },
        { text: 'LOCAÇÃO ', bold: true },
        { text: 'destina-se restritamente ao uso do imóvel para fins ' },
        { text: 'residenciais', bold: true },
        { text: ', estando proibido o ' },
        { text: 'LOCATÁRIO(A) ', bold: true },
        { text: 'sublocá-lo, cedê-lo, transferi-lo ou usá-lo de forma diferente do previsto, salvo autorização expressa por escrito do ' },
        { text: 'LOCADOR(A)', bold: true },
        { text: '.' }
    ]);

    // CLÁUSULA 2
    xml += createClauseTitle('CLÁUSULA 2ª – PRAZO DA LOCAÇÃO');
    xml += createClauseText([
        { text: '1. A presente locação terá a validade de ' },
        { text: dados.contrato.prazo_locacao + ' (' + mesesExtenso + ') meses', bold: true },
        { text: ', a iniciar-se no dia ' },
        { text: formatarData(dados.contrato.data_entrada), bold: true },
        { text: ' e findar-se no dia ' },
        { text: formatarData(dataFim.toISOString().split('T')[0]), bold: true },
        { text: ', data a qual o imóvel deverá ser devolvido nas condições previstas na ' },
        { text: 'cláusula 7ª', bold: true },
        { text: ', efetivando-se com a entrega das chaves, independentemente de aviso ou qualquer outra medida judicial ou extrajudicial.' }
    ]);

    // CLÁUSULA 3
    xml += createClauseTitle('CLÁUSULA 3ª – VALOR DO ALUGUEL');
    xml += createClauseText([
        { text: '2. 1. Como aluguel mensal, o ' },
        { text: 'LOCATÁRIO(A) ', bold: true },
        { text: 'se obrigará a pagar o valor de ' },
        { text: 'R$ ' + dados.contrato.valor_aluguel + ' ', bold: true },
        { text: ', com vencimento sempre no dia ' },
        { text: dados.contrato.data_pagamento, bold: true },
        { text: ' de cada mês.' }
    ]);

    xml += createWordParagraph(
        '2. Fica obrigada o LOCADOR(A), ou seu procurador, a emitir recibo da quantia paga, relacionando pormenorizadamente todos os valores oriundos de juros, ou outra despesa.',
        { alignment: 'both', spacing: '200' }
    );

    xml += createWordParagraph(
        '3. O valor do primeiro mês será calculado de forma proporcional aos dias de uso, conforme data de entrada estabelecida.',
        { alignment: 'both', spacing: '200' }
    );

    xml += createClauseText([
        { text: '4. Emitir-se-á tal recibo, desde que haja a apresentação pelo ' },
        { text: 'LOCATÁRIO(A)', bold: true },
        { text: ', dos comprovantes de todas as despesas do imóvel devidamente quitadas.' }
    ]);

    // CLÁUSULA 4 - CAUÇÃO OU SEGURO
    const tituloClausula4 = dados.contrato.tipo_seguranca === 'caucao'
        ? 'CLÁUSULA 4ª – CAUÇÃO'
        : 'CLÁUSULA 4ª – SEGURO LOCATÍCIO';
    xml += createClauseTitle(tituloClausula4);

    if (dados.contrato.tipo_seguranca === 'caucao') {
        xml += createClauseText([
            { text: '3. Fica estabelecido o valor de ' },
            { text: 'R$ ' + dados.contrato.caucao_valor + ' ', bold: true },
            { text: '(equivalente a ' + caucaoMesesExtenso + ' meses de aluguel), ', bold: true },
            { text: 'a título de ' },
            { text: 'CAUÇÃO', bold: true },
            { text: ', sendo este pago da forma acordada entre as partes.' }
        ]);
    } else {
        xml += createClauseText([
            { text: '3. Fica estabelecida a contratação de ' },
            { text: 'SEGURO LOCATÍCIO ', bold: true },
            { text: 'como garantia da locação, conforme apólice a ser apresentada.' }
        ]);
    }

    // CLÁUSULA 5 - PAGAMENTO
    xml += createClauseTitle('CLÁUSULA 5ª – PAGAMENTO');
    xml += createClauseText([
        { text: '4. 1. Os pagamentos serão efetuados em espécie diretamente à ' },
        { text: 'IMOBILIÁRIA ', bold: true },
        { text: 'através de ' },
        { text: 'depósito, transferência bancária ou PIX ', bold: true },
        { text: 'para a conta corrente digital ' },
        { text: '3463324-0', bold: true },
        { text: ', da agência ' },
        { text: '0001', bold: true },
        { text: ', do ' },
        { text: 'BANCO CORA', bold: true },
        { text: ', chave PIX: ' },
        { text: 'jauaimoveis@yahoo.com.br', bold: true },
        { text: ', nominal à ' },
        { text: 'JAUA IMOVEIS LTDA.ME', bold: true },
        { text: ', cujas parcelas terão sempre o vencimento todo ' },
        { text: 'dia ' + dados.contrato.data_pagamento + ' ', bold: true },
        { text: 'de cada mês, tendo um prazo de tolerância de até ' },
        { text: '05 (cinco) dias ', bold: true },
        { text: 'após o vencimento para efetuar o pagamento do aluguel, mediante aviso por escrito justificando o atraso.' }
    ]);

    xml += createClauseText([
        { text: 'Parágrafo único: ', bold: true },
        { text: 'O primeiro vencimento da locação ocorrerá conforme data estabelecida. O valor referente ao primeiro mês caberá à ' },
        { text: 'JAUA IMOVEIS LTDA.ME ', bold: true },
        { text: '(RE/MAX Litorânea) a título de honorários pelos serviços prestados. A partir do segundo vencimento caberá ao(à) ' },
        { text: 'LOCADOR(A) ', bold: true },
        { text: 'o valor líquido já abatidos 20% referentes à administração.' }
    ]);

    // CLÁUSULA 6 - ATRASO
    xml += createClauseTitle('CLÁUSULA 6ª – DO ATRASO DE PAGAMENTO, MULTA E JUROS APLICÁVEIS');
    xml += createClauseText([
        { text: '5. 1. O ' },
        { text: 'LOCATÁRIO(A)', bold: true },
        { text: ', não vindo a efetuar o pagamento do aluguel até a data estipulada na ' },
        { text: 'cláusula 4.1', bold: true },
        { text: ', fica obrigada a pagar multa de ' },
        { text: '10% (dez por cento) ', bold: true },
        { text: 'sobre o valor do aluguel estipulado neste contrato, bem como juros de mora de ' },
        { text: '1% (um por cento) ', bold: true },
        { text: 'ao mês, mais correção monetária, com prazo máximo de ' },
        { text: '15 (quinze) dias ', bold: true },
        { text: 'para a regularização do pagamento.' }
    ]);

    xml += createClauseText([
        { text: '2. Os pagamentos em atraso após este período poderão ser executados ou protestados, sem comunicado prévio. Em caso de cobrança judicial, o ' },
        { text: 'LOCATÁRIO(A) ', bold: true },
        { text: 'será responsável pelo pagamento das despesas decorrentes dos honorários do advogado e outras provenientes da ação.' }
    ]);

    xml += createClauseText([
        { text: '3. Em caso de atraso igual ou superior a ' },
        { text: '30 (trinta) dias de atraso', bold: true },
        { text: ', este ' },
        { text: 'CONTRATO DE LOCAÇÃO DE IMÓVEL RESIDENCIAL ', bold: true },
        { text: 'estará automaticamente rescindido por motivo de inadimplência, estando o ' },
        { text: 'LOCATÁRIO(A) ', bold: true },
        { text: 'sujeito ao pagamento da multa rescisória estabelecida na ' },
        { text: 'cláusula 12.1', bold: true },
        { text: ', devendo o imóvel ser desocupado imediatamente.' }
    ]);

    // CLÁUSULA 7 - CONTAS
    xml += createClauseTitle('CLÁUSULA 7ª – CONTAS DE ENERGIA, ÁGUA, CONDOMÍNIO E ASSOCIAÇÃO');
    xml += createWordParagraph('7.1 Correrão por conta da Locatária todas as despesas de energia elétrica, água e esgoto, durante o consumo que abrange a vigência deste instrumento.', { alignment: 'both', spacing: '200' });
    xml += createWordParagraph('7.2 Obriga-se o LOCADOR(A) a enviar, por quaisquer meios viáveis, mensalmente, as contas citadas na cláusula anterior, até a data do vencimento, sob pena de arcar com os acréscimos legais (multa e juros) decorrentes.', { alignment: 'both', spacing: '200' });
    xml += createWordParagraph('7.3 Fica o Locatário ciente de que não está autorizado a proceder a transferência de titularidade de qualquer registro, conta ou inscrição vinculado ao imóvel objeto deste contrato.', { alignment: 'both', spacing: '200' });

    // CLÁUSULA 8 - CONDIÇÕES
    xml += createClauseTitle('CLÁUSULA 8ª – CONDIÇÕES DO IMÓVEL, CONSERVAÇÃO, REPAROS E BENFEITORIAS');
    xml += createClauseText([
        { text: '7. 1. O imóvel objeto deste contrato será entregue nas condições descritas no laudo de vistoria, que será realizado na data da entrega das chaves ao ' },
        { text: 'LOCATÁRIO(A)', bold: true },
        { text: ', com instalações elétricas e hidráulicas em perfeito funcionamento, com todos os cômodos e paredes pintados, sendo que portas, portões e acessórios se encontram também em funcionamento correto, devendo a ' },
        { text: 'LOCATÁRIO(A) ', bold: true },
        { text: 'mantê-lo nas mesmas condições em que o está recebendo.' }
    ]);

    xml += createWordParagraph(
        '2. O LOCADOR(A) declara-se integralmente responsável por quaisquer vícios, defeitos ou danos de natureza estrutural existentes ou que venham a se manifestar no imóvel durante a vigência do contrato, desde que estejam expressamente descritos no Laudo de Vistoria inicial ou sejam danos causados pela ação do tempo/natureza, obrigando-se a promover, as suas expensas, todos os reparos necessários, sem ônus ao LOCATÁRIO(A).',
        { alignment: 'both', spacing: '200' }
    );

    xml += createClauseText([
        { text: '3. O reparo de quaisquer danos ao imóvel, não sendo os descritos na clausula 8.2, irão ocorrerão por conta do ' },
        { text: 'LOCATÁRIO(A).', bold: true }
    ]);

    xml += createClauseText([
        { text: '4. Vindo a ser feita benfeitoria devem ser previamente comunicados ao ' },
        { text: 'LOCADOR(A)', bold: true },
        { text: ', e a este faculta aceitá-la ou não, restando ao ' },
        { text: 'LOCATÁRIO(A)', bold: true },
        { text: ', em caso do ' },
        { text: 'LOCADOR(A) ', bold: true },
        { text: 'não a aceitar, modificar o imóvel da maneira que lhe foi entregue.' }
    ]);

    xml += createClauseText([
        { text: '5. As benfeitorias, consertos ou reparos farão parte integrante do imóvel, não assistindo ao ' },
        { text: 'LOCATÁRIO(A) ', bold: true },
        { text: 'o direito de retenção ou indenização sobre as mesmas.' }
    ]);

    // CLÁUSULA 9 - DEVOLUÇÃO
    xml += createClauseTitle('CLÁUSULA 9ª – DEVOLUÇÃO DO IMÓVEL');
    xml += createClauseText([
        { text: '8. 1. O ' },
        { text: 'LOCATÁRIO(A) ', bold: true },
        { text: 'fica obrigada a, no ato da entrega das chaves, devolver o imóvel nas mesmas condições como o está recebendo: limpo e conservado, com as paredes livres de furos, riscos, manchas e danos, em perfeito estado de conservação e asseio, juntamente com todas as instalações de água, energia, portas, janelas, pisos e tudo mais que compõe o ' },
        { text: 'IMÓVEL ', bold: true },
        { text: 'em perfeito estado de conservação e funcionamento conforme laudo de vistoria, salvo caso de benfeitoria aceita pela ' },
        { text: 'LOCADOR(A)', bold: true },
        { text: ', e com todos os tributos e despesas pagas.' }
    ]);

    xml += createClauseText([
        { text: '1. A restituição das chaves ao ' },
        { text: 'LOCADOR(A) ', bold: true },
        { text: 'só poderá ser aceita se o ' },
        { text: 'IMÓVEL ', bold: true },
        { text: 'estiver nas mesmas condições previstas na acima. Se houver necessidade de pintura, obras e reparos, somente após o seu término é que as chaves serão ace, este ficará obrigada a pagar, todas as despesas por danos causados ao imóvel, devendo restituí-lo no estado que lhe foi entregue e que, sobretudo, teve conhecimento no auto de vistoria.' }
    ]);

    // CLÁUSULA 15 - PREFERÊNCIA
    xml += createClauseTitle('CLÁUSULA 15ª – DIREITO DE PREFERÊNCIA');
    xml += createClauseText([
        { text: '14. 1. O ' },
        { text: 'LOCADOR(A)', bold: true },
        { text: ', em qualquer tempo, poderá vender o imóvel, mesmo durante a vigência do contrato de locação e, por via de consequência, ceder os direitos contidos no contrato.' }
    ]);

    xml += createClauseText([
        { text: '2. O ' },
        { text: 'LOCADOR(A) ', bold: true },
        { text: 'deverá notificar o ' },
        { text: 'LOCATÁRIO(A) ', bold: true },
        { text: 'para que este(a) possa exercer seu direito de preferência na aquisição do imóvel, nas mesmas condições em que for oferecido a terceiros.' }
    ]);

    xml += createClauseText([
        { text: '3. Para efetivação da preferência, o ' },
        { text: 'LOCATÁRIO(A) ', bold: true },
        { text: 'deverá responder a notificação, de maneira inequívoca, no prazo de 30 (trinta) dias.' }
    ]);

    xml += createClauseText([
        { text: '3.1 Não havendo interesse na aquisição do imóvel pelo ' },
        { text: 'LOCATÁRIO(A)', bold: true },
        { text: ', este(a) deverá permitir que interessados na compra façam visitas em dias e horários a serem combinados entre ' },
        { text: 'LOCATÁRIO(A)', bold: true },
        { text: ' e ' },
        { text: 'LOCADOR(A)', bold: true },
        { text: '.' }
    ]);

    // CLÁUSULA 16 - DISPOSIÇÕES GERAIS
    xml += createClauseTitle('CLÁUSULA 16ª – DISPOSIÇÕES GERAIS');
    xml += createClauseText([
        { text: '15. 1. As ' },
        { text: 'PARTES ', bold: true },
        { text: 'integrantes deste contrato ficam desde já acordadas a se comunicarem somente por escrito, através de qualquer meio admitido em Direito. Na ausência de qualquer das partes, as mesmas se comprometem desde já a deixarem nomeados procuradores, responsáveis para tal fim.' }
    ]);

    xml += createWordParagraph(
        '2. Os herdeiros, sucessores ou cessionários de ambas as partes se obrigam desde já ao inteiro teor deste contrato.',
        { alignment: 'both', spacing: '200' }
    );

    xml += createClauseText([
        { text: '3. Nenhuma das ' },
        { text: 'PARTES ', bold: true },
        { text: 'poderá ceder ou transferir os direitos e/ou as obrigações deste contrato, bem como sublocar, arrendar, emprestar, no todo ou em parte, a terceiros, sem prévia e expressa anuência da outra ' },
        { text: 'PARTE.', bold: true }
    ]);

    xml += createClauseText([
        { text: '4. Quaisquer alterações nas condições contratadas somente serão efetivadas através de termo aditivo que, uma vez assinado pelas ' },
        { text: 'PARTES', bold: true },
        { text: ', passa a fazer parte integrante do presente contrato.' }
    ]);

    xml += createClauseText([
        { text: '5. As atribuições e obrigações contratuais das partes JAUÁ IMÓVEIS E EMPREENDIMENTOS LTDA. (RE/MAX Litorânea), doravante denominada ' },
        { text: 'IMOBILIÁRIA', bold: true },
        { text: ', assim como, ' },
        { text: dados.corretor.nome, bold: true },
        { text: ', doravante denominado ' },
        { text: 'CORRETOR', bold: true },
        { text: ' ambos já qualificados no presente instrumento, somente responderão pelas tratativas contratuais até o ato de conclusão da presente contratação.' }
    ]);

    xml += createWordParagraph(
        '6. São partes integrantes do presente instrumento de contrato de locação, os anexos a seguir destacados:',
        { alignment: 'both', spacing: '200' }
    );
    xml += createWordParagraph('- Termos e Condições Gerais dos Serviços CREDPAGO;', { alignment: 'both', spacing: '200' });
    xml += createWordParagraph('- Termo de vistoria de início da locação;', { alignment: 'both', spacing: '200' });
    xml += createWordParagraph('7. O presente contrato passa a vigorar entre as partes a partir do ato da assinatura.', { alignment: 'both', spacing: '200' });

    // CLÁUSULA 17 - FORO
    xml += createClauseTitle('CLÁUSULA 17ª – FORO CONTRATUAL');
    xml += createWordParagraph(
        '16. 1. Fica eleito o Foro da Comarca de Camaçari, no Estado da Bahia, para dirimir quaisquer dúvidas e/ou controvérsias oriundas do presente instrumento.',
        { alignment: 'both', spacing: '200' }
    );

    // CLÁUSULA 18 - ASSINATURA
    xml += createClauseTitle('CLÁUSULA 18ª – FORMA DE ASSINATURA');
    xml += createClauseText([
        { text: '17. 1. As ' },
        { text: 'PARTES ', bold: true },
        { text: 'aceitam integralmente que as assinaturas do presente instrumento poderão ser realizadas na forma eletrônica nos termos do parágrafo 2º do artigo 10, da MP 2.200-2/2001. Em caso de assinatura em documento físico, as PARTES firmarão o presente contrato em ' },
        { text: '03 (três) vias ', bold: true },
        { text: 'de igual teor e forma. Em ambos os casos as ' },
        { text: 'PARTES ', bold: true },
        { text: 'assinam na presença de duas testemunhas.' }
    ]);

    // Local e data
    xml += createWordParagraph('', { spacing: '400' });
    xml += createWordParagraph('Camaçari/BA, ___ de _____________ de _______', { alignment: 'center', spacing: '600' });

    // Assinaturas
    xml += createWordParagraph('___________________________________________', { alignment: 'center', spacing: '100' });
    xml += createWordParagraph('LOCADOR(A)', { alignment: 'center', spacing: '100' });
    xml += createWordParagraph(dados.locador.nome, { bold: true, alignment: 'center', spacing: '600' });

    xml += createWordParagraph('___________________________________________', { alignment: 'center', spacing: '100' });
    xml += createWordParagraph('CORRETOR – CRECI/BA ' + dados.corretor.creci, { alignment: 'center', spacing: '100' });
    xml += createWordParagraph(dados.corretor.nome, { bold: true, alignment: 'center', spacing: '600' });

    xml += createWordParagraph('___________________________________________', { alignment: 'center', spacing: '100' });
    xml += createWordParagraph('LOCATÁRIO(A)', { alignment: 'center', spacing: '100' });
    xml += createWordParagraph(dados.locatario.nome, { bold: true, alignment: 'center', spacing: '0' });

    return xml;
}

// Função principal para gerar o contrato usando o template
async function gerarContratoAluguel() {
    if (!validarCamposAluguel()) {
        alert('Por favor, preencha todos os campos obrigatórios.');
        return;
    }

    mostrarLoading();

    try {
        // Verificar disponibilidade do JSZip
        if (typeof JSZip === 'undefined') {
            throw new Error('Biblioteca JSZip não está disponível.');
        }

        const dados = coletarDadosAluguel();

        // Carregar o template MODELO.docx
        const response = await fetch('https://drive.google.com/uc?export=download&id=1IjDWwXtAhNjzyULF-j0Us7A4LytFSZUq');
        if (!response.ok) {
            throw new Error('Não foi possível carregar o template MODELO.docx!');
        }

        const templateBuffer = await response.arrayBuffer();

        // Criar instância JSZip e carregar o template
        const zip = new JSZip();
        await zip.loadAsync(templateBuffer);

        // Ler o document.xml atual
        let documentXml = await zip.file('word/document.xml').async('text');

        // Gerar o conteúdo do contrato em XML
        const contractContentXml = generateContractXml(dados);

        // Extrair o namespace do documento original
        const bodyMatch = documentXml.match(/<w:body>([\s\S]*?)<\/w:body>/);
        if (!bodyMatch) {
            throw new Error('Estrutura do documento template inválida');
        }

        // Substituir o conteúdo do body mantendo a estrutura
        // Preservar o sectPr (configurações de seção) do final
        const sectPrMatch = documentXml.match(/(<w:sectPr[\s\S]*?<\/w:sectPr>)/);
        const sectPr = sectPrMatch ? sectPrMatch[1] : '';

        // Criar novo document.xml
        const newDocumentXml = documentXml.replace(
            /<w:body>[\s\S]*?<\/w:body>/,
            `<w:body>${contractContentXml}${sectPr}</w:body>`
        );

        // Atualizar o arquivo no ZIP
        zip.file('word/document.xml', newDocumentXml);

        // Gerar o novo arquivo
        const newDocBuffer = await zip.generateAsync({ type: 'arraybuffer' });

        // Criar blob e fazer download
        const blob = new Blob([newDocBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

        // Nome do arquivo: CONTRATO + NOME LOCADOR + NOME LOCATÁRIO.docx
        const nomeLocador = dados.locador.nome.split(' ')[0];
        const nomeLocatario = dados.locatario.nome.split(' ')[0];
        const fileName = `CONTRATO ${nomeLocador} ${nomeLocatario}.docx`;

        // Download
        if (window.saveAs) {
            window.saveAs(blob, fileName);
        } else {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        ocultarLoading();
        alert("Contrato gerado com sucesso!");

    } catch (error) {
        console.error("Erro detalhado:", error);
        ocultarLoading();
        alert("Erro ao gerar contrato: " + error.message);
    }
}