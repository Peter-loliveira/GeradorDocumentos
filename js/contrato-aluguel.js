// contrato-aluguel.js - Gerador de Contrato de Locação

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

async function gerarContratoAluguel() {
    if (!validarCamposAluguel()) {
        alert('Por favor, preencha todos os campos obrigatórios.');
        return;
    }

    mostrarLoading();

    try {
        // Verificar disponibilidade das bibliotecas
        if (typeof window.docx === 'undefined') {
            throw new Error('Biblioteca docx não está disponível. Verifique sua conexão com a internet ou se o script foi bloqueado.');
        }

        // Acessar docx através do objeto window
        const docxLib = window.docx;
        console.log('docxLib carregado:', docxLib);
        
        const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } = docxLib;
        
        const dados = coletarDadosAluguel();
        
        const dataEntrada = new Date(dados.contrato.data_entrada);
        const dataFim = new Date(dataEntrada);
        dataFim.setMonth(dataFim.getMonth() + parseInt(dados.contrato.prazo_locacao));
        
        const mesesExtenso = numeroPorExtenso(parseInt(dados.contrato.prazo_locacao));
        const caucaoMesesExtenso = dados.contrato.caucao_meses ? numeroPorExtenso(parseInt(dados.contrato.caucao_meses)) : '';

        const children = [];
        
        // Título
        children.push(
            new Paragraph({
                text: "CONTRATO DE LOCAÇÃO DE IMÓVEL RESIDENCIAL",
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 }
            })
        );

        // Função auxiliar para criar parágrafos das partes
        const criarParagrafoParte = (titulo, conteudo) => {
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: titulo, bold: true, size: 24 })],
                    spacing: { after: 200 }
                }),
                new Paragraph({
                    children: conteudo,
                    spacing: { after: 300 },
                    alignment: AlignmentType.JUSTIFIED
                })
            );
        };

        // LOCADOR
        criarParagrafoParte("LOCADOR(A)", [
            new TextRun({ text: dados.locador.nome + ", ", bold: true }),
            new TextRun({ text: dados.locador.nacionalidade + ", " + dados.locador.estado_civil.toLowerCase() + ", portador(a) do " }),
            new TextRun({ text: "RG nº " + dados.locador.rg + ", ", bold: true }),
            new TextRun({ text: "emitido pela " + dados.locador.rg_orgao + ", inscrito(a) no " }),
            new TextRun({ text: "CPF sob nº " + dados.locador.cpf + " ", bold: true }),
            new TextRun({ text: "residente e domiciliado(a) na " + dados.locador.endereco + ", Telefone: " + dados.locador.telefone + ", E-mail: " + dados.locador.email + ", representado(a) neste ato pelo(a) corretor(a) " }),
            new TextRun({ text: dados.corretor.nome + ", ", bold: true }),
            new TextRun({ text: "corretor(a) de imóveis, inscrito(a) no " }),
            new TextRun({ text: "CPF sob o nº " + dados.corretor.cpf + ", ", bold: true }),
            new TextRun({ text: "registrado(a) no " }),
            new TextRun({ text: "CRECI-BA sob o nº " + dados.corretor.creci + ", ", bold: true }),
            new TextRun({ text: "telefone de contato +55 (71) 999441701, e-mail peteroliveira@remax.com.br." })
        ]);

        // LOCATÁRIO
        criarParagrafoParte("LOCATÁRIO(A)", [
            new TextRun({ text: dados.locatario.nome + ", ", bold: true }),
            new TextRun({ text: dados.locatario.nacionalidade + ", " + dados.locatario.estado_civil.toLowerCase() + ", portador(a) do " }),
            new TextRun({ text: "RG nº " + dados.locatario.rg + ", ", bold: true }),
            new TextRun({ text: "emitido pela " + dados.locatario.rg_orgao + ", inscrito(a) no " }),
            new TextRun({ text: "CPF sob nº " + dados.locatario.cpf + ", ", bold: true }),
            new TextRun({ text: "residente e domiciliado(a) na " + dados.locatario.endereco + ", telefone " + dados.locatario.telefone + ", e-mail " + dados.locatario.email + "." })
        ]);

        // IMOBILIÁRIA
        criarParagrafoParte("IMOBILIÁRIA", [
            new TextRun({ text: "JAUÁ IMÓVEIS E EMPREENDIMENTOS LTDA. (RE/MAX Litorânea), inscrita no CNPJ sob o nº 07.788.314.0001-40, registrada no CRECI-BA sob o nº 1101 PJ, com sede na Rua Direta de Jaua, Loja 4, Jaua, Camaçari, Bahia, CEP: 42828-576, com telefone de contato +55 (71) 3672-1664 e e-mail litoranea@remax.com.br." })
        ]);

        // CORRETOR
        criarParagrafoParte("CORRETOR", [
            new TextRun({ text: dados.corretor.nome + ", ", bold: true }),
            new TextRun({ text: "corretor de imóveis, inscrito no CPF sob o nº " + dados.corretor.cpf + ", registrada no CRECI-BA sob o nº " + dados.corretor.creci + ", telefone de contato +55 (71) 9-9944-1701, e-mail peteroliveira@remax.com.br." })
        ]);

        // Introdução
        children.push(
            new Paragraph({
                children: [
                    new TextRun({ text: "As partes acima qualificadas estabelecem entre si o presente " }),
                    new TextRun({ text: "CONTRATO DE LOCAÇÃO DE IMÓVEL RESIDENCIAL ", bold: true }),
                    new TextRun({ text: "mediante as condições e cláusulas seguintes:" })
                ],
                spacing: { after: 400 },
                alignment: AlignmentType.JUSTIFIED
            })
        );

        // Função para criar cláusulas
        const criarClausula = (titulo, paragrafos) => {
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: titulo, bold: true, size: 24 })],
                    spacing: { after: 200 }
                })
            );
            paragrafos.forEach(p => {
                children.push(
                    new Paragraph({
                        children: p,
                        spacing: { after: 200 },
                        alignment: AlignmentType.JUSTIFIED
                    })
                );
            });
            children.push(new Paragraph({ spacing: { after: 200 } }));
        };

        // CLÁUSULA 1
        criarClausula("CLÁUSULA 1ª – OBJETO DO CONTRATO", [
            [
                new TextRun({ text: "1. O presente instrumento tem por " }),
                new TextRun({ text: "OBJETO ", bold: true }),
                new TextRun({ text: "o imóvel do tipo " }),
                new TextRun({ text: dados.contrato.tipo_imovel.toLowerCase() + " ", bold: true }),
                new TextRun({ text: "de propriedade do " }),
                new TextRun({ text: "LOCADOR(A)", bold: true }),
                new TextRun({ text: ", situado na " + dados.contrato.imovel_endereco + ", " + dados.contrato.imovel_descricao + "." })
            ],
            [
                new TextRun({ text: "Parágrafo primeiro: ", bold: true }),
                new TextRun({ text: "Quando do início da locação será lavrado laudo de vistoria no qual constará, pormenorizadamente, a descrição da quantidade, qualidade e espécies de móveis e utensílios existentes, bem como do estado de conservação do imóvel, suas instalações hidráulicas e elétricas." })
            ],
            [
                new TextRun({ text: "Parágrafo segundo: ", bold: true }),
                new TextRun({ text: "A presente " }),
                new TextRun({ text: "LOCAÇÃO ", bold: true }),
                new TextRun({ text: "destina-se restritamente ao uso do imóvel para fins " }),
                new TextRun({ text: "residenciais", bold: true }),
                new TextRun({ text: ", estando proibido o " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "sublocá-lo, cedê-lo, transferi-lo ou usá-lo de forma diferente do previsto, salvo autorização expressa por escrito do " }),
                new TextRun({ text: "LOCADOR(A)", bold: true }),
                new TextRun({ text: "." })
            ]
        ]);

        // CLÁUSULA 2
        criarClausula("CLÁUSULA 2ª – PRAZO DA LOCAÇÃO", [
            [
                new TextRun({ text: "1. A presente locação terá a validade de " }),
                new TextRun({ text: dados.contrato.prazo_locacao + " (" + mesesExtenso + ") meses", bold: true }),
                new TextRun({ text: ", a iniciar-se no dia " }),
                new TextRun({ text: formatarData(dados.contrato.data_entrada), bold: true }),
                new TextRun({ text: " e findar-se no dia " }),
                new TextRun({ text: formatarData(dataFim.toISOString().split('T')[0]), bold: true }),
                new TextRun({ text: ", data a qual o imóvel deverá ser devolvido nas condições previstas na " }),
                new TextRun({ text: "cláusula 7ª", bold: true }),
                new TextRun({ text: ", efetivando-se com a entrega das chaves, independentemente de aviso ou qualquer outra medida judicial ou extrajudicial." })
            ]
        ]);

        // CLÁUSULA 3
        criarClausula("CLÁUSULA 3ª – VALOR DO ALUGUEL", [
            [
                new TextRun({ text: "2. 1. Como aluguel mensal, o " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "se obrigará a pagar o valor de " }),
                new TextRun({ text: "R$ " + dados.contrato.valor_aluguel + " ", bold: true }),
                new TextRun({ text: ", com vencimento sempre no dia " }),
                new TextRun({ text: dados.contrato.data_pagamento, bold: true }),
                new TextRun({ text: " de cada mês." })
            ],
            [
                new TextRun({ text: "2. Fica obrigada o " }),
                new TextRun({ text: "LOCADOR(A)", bold: true }),
                new TextRun({ text: ", ou seu procurador, a emitir recibo da quantia paga, relacionando pormenorizadamente todos os valores oriundos de juros, ou outra despesa." })
            ],
            [
                new TextRun({ text: "3. O valor do primeiro mês será calculado de forma proporcional aos dias de uso, conforme data de entrada estabelecida." })
            ],
            [
                new TextRun({ text: "4. Emitir-se-á tal recibo, desde que haja a apresentação pelo " }),
                new TextRun({ text: "LOCATÁRIO(A)", bold: true }),
                new TextRun({ text: ", dos comprovantes de todas as despesas do imóvel devidamente quitadas." })
            ]
        ]);

        // CLÁUSULA 4 - CAUÇÃO
        const clausula4Paragrafos = dados.contrato.tipo_seguranca === 'caucao' ? [
            [
                new TextRun({ text: "3. Fica estabelecido o valor de " }),
                new TextRun({ text: "R$ " + dados.contrato.caucao_valor + " ", bold: true }),
                new TextRun({ text: "(equivalente a" + caucaoMesesExtenso + " meses de aluguel), ", bold: true }),
                new TextRun({ text: "a título de " }),
                new TextRun({ text: "CAUÇÃO", bold: true }),
                new TextRun({ text: ", sendo este pago da forma acordada entre as partes." })
            ]
        ] : [
            [
                new TextRun({ text: "3. Fica estabelecida a contratação de " }),
                new TextRun({ text: "SEGURO LOCATÍCIO ", bold: true }),
                new TextRun({ text: "como garantia da locação, conforme apólice a ser apresentada." })
            ]
        ];
        criarClausula("CLÁUSULA 4ª – CAUÇÃO", clausula4Paragrafos);

        // CLÁUSULA 5 - PAGAMENTO
        criarClausula("CLÁUSULA 5ª – PAGAMENTO", [
            [
                new TextRun({ text: "4. 1. Os pagamentos serão efetuados em espécie diretamente à " }),
                new TextRun({ text: "IMOBILIÁRIA ", bold: true }),
                new TextRun({ text: "através de " }),
                new TextRun({ text: "depósito, transferência bancária ou PIX ", bold: true }),
                new TextRun({ text: "para a conta corrente digital " }),
                new TextRun({ text: "3463324-0", bold: true }),
                new TextRun({ text: ", da agência " }),
                new TextRun({ text: "0001", bold: true }),
                new TextRun({ text: ", do " }),
                new TextRun({ text: "BANCO CORA", bold: true }),
                new TextRun({ text: ", chave PIX: " }),
                new TextRun({ text: "jauaimoveis@yahoo.com.br", bold: true }),
                new TextRun({ text: ", nominal à " }),
                new TextRun({ text: "JAUA IMOVEIS LTDA.ME", bold: true }),
                new TextRun({ text: ", cujas parcelas terão sempre o vencimento todo " }),
                new TextRun({ text: "dia " + dados.contrato.data_pagamento + " ", bold: true }),
                new TextRun({ text: "de cada mês, tendo um prazo de tolerância de até " }),
                new TextRun({ text: "05 (cinco) dias ", bold: true }),
                new TextRun({ text: "após o vencimento para efetuar o pagamento do aluguel, mediante aviso por escrito justificando o atraso." })
            ],
            [
                new TextRun({ text: "Parágrafo único: ", bold: true }),
                new TextRun({ text: "O primeiro vencimento da locação ocorrerá conforme data estabelecida. O valor referente ao primeiro mês caberá à " }),
                new TextRun({ text: "JAUA IMOVEIS LTDA.ME ", bold: true }),
                new TextRun({ text: "(RE/MAX Litorânea) a título de honorários pelos serviços prestados. A partir do segundo vencimento caberá ao(à) " }),
                new TextRun({ text: "LOCADOR(A) ", bold: true }),
                new TextRun({ text: "o valor líquido já abatidos 20% referentes à administração." })
            ]
        ]);

        // CLÁUSULA 6 - ATRASO
        criarClausula("CLÁUSULA 5ª – DO ATRASO DE PAGAMENTO, MULTA E JUROS APLICÁVEIS", [
            [
                new TextRun({ text: "5. 1. O " }),
                new TextRun({ text: "LOCATÁRIO(A)", bold: true }),
                new TextRun({ text: ", não vindo a efetuar o pagamento do aluguel até a data estipulada na " }),
                new TextRun({ text: "cláusula 4.1", bold: true }),
                new TextRun({ text: ", fica obrigada a pagar multa de " }),
                new TextRun({ text: "10% (dez por cento) ", bold: true }),
                new TextRun({ text: "sobre o valor do aluguel estipulado neste contrato, bem como juros de mora de " }),
                new TextRun({ text: "1% (um por cento) ", bold: true }),
                new TextRun({ text: "ao mês, mais correção monetária, com prazo máximo de " }),
                new TextRun({ text: "15 (quinze) dias ", bold: true }),
                new TextRun({ text: "para a regularização do pagamento." })
            ],
            [
                new TextRun({ text: "2. Os pagamentos em atraso após este período poderão ser executados ou protestados, sem comunicado prévio. Em caso de cobrança judicial, o " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "será responsável pelo pagamento das despesas decorrentes dos honorários do advogado e outras provenientes da ação." })
            ],
            [
                new TextRun({ text: "3. Em caso de atraso igual ou superior a " }),
                new TextRun({ text: "30 (trinta) dias de atraso", bold: true }),
                new TextRun({ text: ", este " }),
                new TextRun({ text: "CONTRATO DE LOCAÇÃO DE IMÓVEL RESIDENCIAL ", bold: true }),
                new TextRun({ text: "estará automaticamente rescindido por motivo de inadimplência, estando o " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "sujeito ao pagamento da multa rescisória estabelecida na " }),
                new TextRun({ text: "cláusula 12.1", bold: true }),
                new TextRun({ text: ", devendo o imóvel ser desocupado imediatamente." })
            ]
        ]);

        // CLÁUSULA 7 - CONTAS
        criarClausula("CLÁUSULA 7ª – CONTAS DE ENERGIA, ÁGUA, CONDOMÍNIO E ASSOCIAÇÃO", [
            [new TextRun({ text: "7.1 Correrão por conta da Locatária todas as despesas de energia elétrica, água e esgoto, durante o consumo que abrange a vigência deste instrumento." })],
            [new TextRun({ text: "7.2 Obriga-se o LOCADOR(A) a enviar, por quaisquer meios viáveis, mensalmente, as contas citadas na cláusula anterior, até a data do vencimento, sob pena de arcar com os acréscimos legais (multa e juros) decorrentes." })],
            [new TextRun({ text: "7.3 Fica o Locatário ciente de que não está autorizado a proceder a transferência de titularidade de qualquer registro, conta ou inscrição vinculado ao imóvel objeto deste contrato." })]
        ]);

        // CLÁUSULA 8 - CONDIÇÕES
        criarClausula("CLÁUSULA 8ª – CONDIÇÕES DO IMÓVEL, CONSERVAÇÃO, REPAROS E BENFEITORIAS", [
            [
                new TextRun({ text: "7. 1. O imóvel objeto deste contrato será entregue nas condições descritas no laudo de vistoria, que será realizado na data da entrega das chaves ao " }),
                new TextRun({ text: "LOCATÁRIO(A)", bold: true }),
                new TextRun({ text: ", com instalações elétricas e hidráulicas em perfeito funcionamento, com todos os cômodos e paredes pintados, sendo que portas, portões e acessórios se encontram também em funcionamento correto, devendo a " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "mantê-lo nas mesmas condições em que o está recebendo." })
            ],
            [
                new TextRun({ text: "2. O LOCADOR(A) declara-se integralmente responsável por quaisquer vícios, defeitos ou danos de natureza estrutural existentes ou que venham a se manifestar no imóvel durante a vigência do contrato, desde que estejam expressamente descritos no Laudo de Vistoria inicial ou sejam danos causados pela ação do tempo/natureza, obrigando-se a promover, as suas expensas, todos os reparos necessários, sem ônus ao " }),
                new TextRun({ text: "LOCATÁRIO(A).", bold: true })
            ],
            [
                new TextRun({ text: "3. O reparo de quaisquer danos ao imóvel, não sendo os descritos na clausula 8.2, irão ocorrerão por conta do " }),
                new TextRun({ text: "LOCATÁRIO(A).", bold: true })
            ],
            [
                new TextRun({ text: "4. Vindo a ser feita benfeitoria devem ser previamente comunicados ao " }),
                new TextRun({ text: "LOCADOR(A)", bold: true }),
                new TextRun({ text: ", e a este faculta aceitá-la ou não, restando ao " }),
                new TextRun({ text: "LOCATÁRIO(A)", bold: true }),
                new TextRun({ text: ", em caso do " }),
                new TextRun({ text: "LOCADOR(A) ", bold: true }),
                new TextRun({ text: "não a aceitar, modificar o imóvel da maneira que lhe foi entregue." })
            ],
            [
                new TextRun({ text: "5. As benfeitorias, consertos ou reparos farão parte integrante do imóvel, não assistindo ao " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "o direito de retenção ou indenização sobre as mesmas." })
            ]
        ]);

        // CLÁUSULA 9 - DEVOLUÇÃO
        criarClausula("CLÁUSULA 9ª – DEVOLUÇÃO DO IMÓVEL", [
            [
                new TextRun({ text: "8. 1. O " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "fica obrigada a, no ato da entrega das chaves, devolver o imóvel nas mesmas condições como o está recebendo: limpo e conservado, com as paredes livres de furos, riscos, manchas e danos, em perfeito estado de conservação e asseio, juntamente com todas as instalações de água, energia, portas, janelas, pisos e tudo mais que compõe o " }),
                new TextRun({ text: "IMÓVEL ", bold: true }),
                new TextRun({ text: "em perfeito estado de conservação e funcionamento conforme laudo de vistoria, salvo caso de benfeitoria aceita pela " }),
                new TextRun({ text: "LOCADOR(A)", bold: true }),
                new TextRun({ text: ", e com todos os tributos e despesas pagas." })
            ],
            [
                new TextRun({ text: "1. A restituição das chaves ao " }),
                new TextRun({ text: "LOCADOR(A) ", bold: true }),
                new TextRun({ text: "só poderá ser aceita se o " }),
                new TextRun({ text: "IMÓVEL ", bold: true }),
                new TextRun({ text: "estiver nas mesmas condições previstas na acima. Se houver necessidade de pintura, obras e reparos, somente após o seu término é que as chaves serão aceitas pelo " }),
                new TextRun({ text: "LOCADOR(A)", bold: true }),
                new TextRun({ text: ", devendo o " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "providenciar as devidas providências." })
            ]
        ]);

        // CLÁUSULA 10 - VISTORIA
        criarClausula("CLÁUSULA 10ª – VISTORIA", [
            [
                new TextRun({ text: "9. 1. A fim de verificar o exato cumprimento das obrigações assumidas neste contrato, reserva-se ao " }),
                new TextRun({ text: "LOCADOR(A) ", bold: true }),
                new TextRun({ text: "o direito de vistoriar o " }),
                new TextRun({ text: "IMÓVEL ", bold: true }),
                new TextRun({ text: "pessoalmente ou por seu representante, a qualquer tempo, mediante agendamento prévio com o " }),
                new TextRun({ text: "LOCATÁRIO(A).", bold: true })
            ]
        ]);

        // CLÁUSULA 11 - RENOVAÇÃO
        criarClausula("CLÁUSULA 11ª – RENOVAÇÃO", [
            [
                new TextRun({ text: "10. 1. O " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "terá a preferência na renovação do " }),
                new TextRun({ text: "CONTRATO DE LOCAÇÃO DE IMÓVEL RESIDENCIAL", bold: true }),
                new TextRun({ text: ", devendo comunicar por escrito ao " }),
                new TextRun({ text: "LOCADOR(A) ", bold: true }),
                new TextRun({ text: "sua intenção de renovação com no mínimo " }),
                new TextRun({ text: "60 (sessenta) dias ", bold: true }),
                new TextRun({ text: "de antecedência ao vencimento do presente contrato." })
            ],
            [
                new TextRun({ text: "2. A renovação contratual, entretanto, não será automática, podendo o " }),
                new TextRun({ text: "LOCADOR(A) ", bold: true }),
                new TextRun({ text: "optar pela não renovação. Neste caso, o " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "deverá desocupar o " }),
                new TextRun({ text: "IMÓVEL ", bold: true }),
                new TextRun({ text: "no último dia de vigência do presente " }),
                new TextRun({ text: "CONTRATO DE LOCAÇÃO DE IMÓVEL RESIDENCIAL.", bold: true })
            ]
        ]);

        // CLÁUSULA 12 - REAJUSTE
        criarClausula("CLÁUSULA 12ª – REAJUSTE DO ALUGUEL", [
            [
                new TextRun({ text: "11. 1. O valor do aluguel será reajustado a cada período de 12(doze) meses, tendo como base, a variação dos índices IGPM, IGP, IPC, etc., ocorrido no período anual, ou em sua falta ou extinção, será substituída pelo maior índice oficial vigente." })
            ],
            [
                new TextRun({ text: "Parágrafo único: ", bold: true }),
                new TextRun({ text: "Em caso de falta deste índice, o reajustamento do aluguel terá por base a média da variação dos índices inflacionários do ano corrente ao da execução do aluguel, até o primeiro dia anterior ao pagamento de todos os valores devidos. Ocorrendo alguma mudança no âmbito governamental, todos os valores agregados ao aluguel, bem como o próprio aluguel, serão revistos pelas partes." })
            ]
        ]);

        // CLÁUSULA 13 - RESCISÃO
        criarClausula("CLÁUSULA 13ª – RESCISÃO E MULTAS", [
            [
                new TextRun({ text: "12. 1. No caso de rescisão por descumprimento de qualquer uma das cláusulas deste " }),
                new TextRun({ text: "CONTRATO DE LOCAÇÃO DE IMÓVEL", bold: true }),
                new TextRun({ text: ", o responsável pela rescisão ficará sujeito ao pagamento da multa, equivalente à " }),
                new TextRun({ text: "03 (três) aluguéis proporcionais).", bold: true })
            ],
            [
                new TextRun({ text: "2. Ocorrerá a rescisão do presente contrato, independentemente de qualquer comunicação prévia ou indenização por parte do " }),
                new TextRun({ text: "LOCATÁRIO(A)", bold: true }),
                new TextRun({ text: ", quando:" })
            ],
            [
                new TextRun({ text: "2.1 Ocorrendo qualquer sinistro, incêndio ou algo que venha a impossibilitar a posse do imóvel, independente de dolo ou culpa do " }),
                new TextRun({ text: "LOCATÁRIO(A)", bold: true }),
                new TextRun({ text: "; bem como quaisquer outras hipóteses que maculem o imóvel de vício e impossibilite sua posse;" })
            ],
            [
                new TextRun({ text: "2.2 Em hipótese de desapropriação do imóvel alugado." })
            ],
            [
                new TextRun({ text: "3. Poderá também o presente instrumento ser rescindido, sem gerar direito a indenização ou qualquer ônus para o " }),
                new TextRun({ text: "LOCADOR(A)", bold: true }),
                new TextRun({ text: ", caso o imóvel seja utilizado de forma diversa da locação residencial, sem prejuízo da obrigação do " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "de efetuar o pagamento das multas e despesas previstas na " }),
                new TextRun({ text: "cláusula 5ª", bold: true }),
                new TextRun({ text: ", salvo autorização expressa do " }),
                new TextRun({ text: "LOCADOR(A).", bold: true })
            ],
            [
                new TextRun({ text: "4. Caso o motivo da rescisão seja a antecipação da entrega do " }),
                new TextRun({ text: "IMÓVEL", bold: true }),
                new TextRun({ text: ", antes de completar 12 (doze) meses de locação, a " }),
                new TextRun({ text: "PARTE SOLICITANTE ", bold: true }),
                new TextRun({ text: "deverá comunicar à " }),
                new TextRun({ text: "OUTRA PARTE ", bold: true }),
                new TextRun({ text: "por escrito no prazo mínimo de " }),
                new TextRun({ text: "30 (trinta) dias ", bold: true }),
                new TextRun({ text: "de antecedência, considerada da data em que o " }),
                new TextRun({ text: "IMÓVEL ", bold: true }),
                new TextRun({ text: "deverá ser desocupado, incidindo à " }),
                new TextRun({ text: "PARTE SOLICITANTE", bold: true }),
                new TextRun({ text: ", entretanto, o pagamento de uma multa equivalente à " }),
                new TextRun({ text: "03 (três) aluguéis ", bold: true }),
                new TextRun({ text: "à " }),
                new TextRun({ text: "OUTRA PARTE.", bold: true })
            ],
            [
                new TextRun({ text: "5. Caso a rescisão venha a acontecer a partir do 13º (décimo terceiro) mês de locação, a " }),
                new TextRun({ text: "PARTE SOLICITANTE ", bold: true }),
                new TextRun({ text: "deverá comunicar à " }),
                new TextRun({ text: "OUTRA PARTE ", bold: true }),
                new TextRun({ text: "por escrito no prazo mínimo de " }),
                new TextRun({ text: "30 (trinta) dias ", bold: true }),
                new TextRun({ text: "de antecedência, neste caso ambas as " }),
                new TextRun({ text: "PARTES ", bold: true }),
                new TextRun({ text: "estão isentas da multa prevista na " }),
                new TextRun({ text: "cláusula 13.4.", bold: true })
            ],
            [
                new TextRun({ text: "6. O pagamento da multa rescisória não exonera o " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "de entregar o " }),
                new TextRun({ text: "IMÓVEL ", bold: true }),
                new TextRun({ text: "nas condições estabelecidas neste contrato." })
            ]
        ]);

        // CLÁUSULA 14 - DANOS
        criarClausula("CLÁUSULA 14ª – DOS DANOS AO IMÓVEL", [
            [
                new TextRun({ text: "13. 1. O " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "se responsabiliza por quaisquer danos que venham a causar ao imóvel, ou aos bens que o guarnecem, devendo restituí-los nas mesmas condições em que o recebeu." })
            ],
            [
                new TextRun({ text: "2. Qualquer acidente que porventura venha a ocorrer no imóvel por culpa ou dolo do " }),
                new TextRun({ text: "LOCATÁRIO(A)", bold: true }),
                new TextRun({ text: ", este ficará obrigada a pagar, todas as despesas por danos causados ao imóvel, devendo restituí-lo no estado que lhe foi entregue e que, sobretudo, teve conhecimento no auto de vistoria." })
            ]
        ]);

        // CLÁUSULA 15 - PREFERÊNCIA
        criarClausula("CLÁUSULA 15ª – DIREITO DE PREFERÊNCIA", [
            [
                new TextRun({ text: "14. 1. O " }),
                new TextRun({ text: "LOCADOR(A)", bold: true }),
                new TextRun({ text: ", em qualquer tempo, poderá vender o imóvel, mesmo durante a vigência do contrato de locação e, por via de consequência, ceder os direitos contidos no contrato." })
            ],
            [
                new TextRun({ text: "2. O " }),
                new TextRun({ text: "LOCADOR(A) ", bold: true }),
                new TextRun({ text: "deverá notificar o " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "para que este(a) possa exercer seu direito de preferência na aquisição do imóvel, nas mesmas condições em que for oferecido a terceiros." })
            ],
            [
                new TextRun({ text: "3. Para efetivação da preferência, o " }),
                new TextRun({ text: "LOCATÁRIO(A) ", bold: true }),
                new TextRun({ text: "deverá responder a notificação, de maneira inequívoca, no prazo de 30 (trinta) dias." })
            ],
            [
                new TextRun({ text: "3.1 Não havendo interesse na aquisição do imóvel pelo " }),
                new TextRun({ text: "LOCATÁRIO(A)", bold: true }),
                new TextRun({ text: ", este(a) deverá permitir que interessados na compra façam visitas em dias e horários a serem combinados entre " }),
                new TextRun({ text: "LOCATÁRIO(A)", bold: true }),
                new TextRun({ text: " e " }),
                new TextRun({ text: "LOCADOR(A)", bold: true }),
                new TextRun({ text: "." })
            ]
        ]);

        // CLÁUSULA 16 - DISPOSIÇÕES GERAIS
        criarClausula("CLÁUSULA 16ª – DISPOSIÇÕES GERAIS", [
            [
                new TextRun({ text: "15. 1. As " }),
                new TextRun({ text: "PARTES ", bold: true }),
                new TextRun({ text: "integrantes deste contrato ficam desde já acordadas a se comunicarem somente por escrito, através de qualquer meio admitido em Direito. Na ausência de qualquer das partes, as mesmas se comprometem desde já a deixarem nomeados procuradores, responsáveis para tal fim." })
            ],
            [
                new TextRun({ text: "2. Os herdeiros, sucessores ou cessionários de ambas as partes se obrigam desde já ao inteiro teor deste contrato." })
            ],
            [
                new TextRun({ text: "3. Nenhuma das " }),
                new TextRun({ text: "PARTES ", bold: true }),
                new TextRun({ text: "poderá ceder ou transferir os direitos e/ou as obrigações deste contrato, bem como sublocar, arrendar, emprestar, no todo ou em parte, a terceiros, sem prévia e expressa anuência da outra " }),
                new TextRun({ text: "PARTE.", bold: true })
            ],
            [
                new TextRun({ text: "4. Quaisquer alterações nas condições contratadas somente serão efetivadas através de termo aditivo que, uma vez assinado pelas " }),
                new TextRun({ text: "PARTES", bold: true }),
                new TextRun({ text: ", passa a fazer parte integrante do presente contrato." })
            ],
            [
                new TextRun({ text: "5. As atribuições e obrigações contratuais das partes JAUÁ IMÓVEIS E EMPREENDIMENTOS LTDA. (RE/MAX Litorânea), doravante denominada " }),
                new TextRun({ text: "IMOBILIÁRIA", bold: true }),
                new TextRun({ text: ", assim como, " }),
                new TextRun({ text: dados.corretor.nome, bold: true }),
                new TextRun({ text: ", doravante denominado " }),
                new TextRun({ text: "CORRETOR", bold: true }),
                new TextRun({ text: " ambos já qualificados no presente instrumento, somente responderão pelas tratativas contratuais até o ato de conclusão da presente contratação." })
            ],
            [
                new TextRun({ text: "6. São partes integrantes do presente instrumento de contrato de locação, os anexos a seguir destacados:" })
            ],
            [
                new TextRun({ text: "- Termos e Condições Gerais dos Serviços CREDPAGO;" })
            ],
            [
                new TextRun({ text: "- Termo de vistoria de início da locação;" })
            ],
            [
                new TextRun({ text: "7. O presente contrato passa a vigorar entre as partes a partir do ato da assinatura." })
            ]
        ]);

        // CLÁUSULA 17 - FORO
        criarClausula("CLÁUSULA 17ª – FORO CONTRATUAL", [
            [
                new TextRun({ text: "16. 1. Fica eleito o Foro da Comarca de Camaçari, no Estado da Bahia, para dirimir quaisquer dúvidas e/ou controvérsias oriundas do presente instrumento." })
            ]
        ]);

        // CLÁUSULA 18 - ASSINATURA
        criarClausula("CLÁUSULA 18ª – FORMA DE ASSINATURA", [
            [
                new TextRun({ text: "17. 1. As " }),
                new TextRun({ text: "PARTES ", bold: true }),
                new TextRun({ text: "aceitam integralmente que as assinaturas do presente instrumento poderão ser realizadas na forma eletrônica nos termos do parágrafo 2º do artigo 10, da MP 2.200-2/2001. Em caso de assinatura em documento físico, as PARTES firmarão o presente contrato em " }),
                new TextRun({ text: "03 (três) vias ", bold: true }),
                new TextRun({ text: "de igual teor e forma. Em ambos os casos as " }),
                new TextRun({ text: "PARTES ", bold: true }),
                new TextRun({ text: "assinam na presença de duas testemunhas." })
            ]
        ]);

        // Local e data
        children.push(
            new Paragraph({ spacing: { after: 400 } }),
            new Paragraph({
                text: "Camaçari/BA, ___ de _____________ de _______",
                alignment: AlignmentType.CENTER,
                spacing: { after: 600 }
            })
        );

        // Assinaturas
        children.push(
            new Paragraph({
                text: "___________________________________________",
                alignment: AlignmentType.CENTER,
                spacing: { before: 400, after: 100 }
            }),
            new Paragraph({
                text: "LOCADOR(A)",
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new Paragraph({
                children: [new TextRun({ text: dados.locador.nome, bold: true })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 600 }
            }),
            new Paragraph({
                text: "___________________________________________",
                alignment: AlignmentType.CENTER,
                spacing: { before: 400, after: 100 }
            }),
            new Paragraph({
                text: "CORRETOR – CRECI/BA " + dados.corretor.creci,
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new Paragraph({
                children: [new TextRun({ text: dados.corretor.nome, bold: true })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 600 }
            }),
            new Paragraph({
                text: "___________________________________________",
                alignment: AlignmentType.CENTER,
                spacing: { before: 400, after: 100 }
            }),
            new Paragraph({
                text: "LOCATÁRIO(A)",
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new Paragraph({
                children: [new TextRun({ text: dados.locatario.nome, bold: true })],
                alignment: AlignmentType.CENTER
            })
        );

        // Criar documento
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            right: 1440,
                            bottom: 1440,
                            left: 1440
                        }
                    }
                },
                children: children
            }]
        });

        // Gerar arquivo
        const blob = await Packer.toBlob(doc);
        const fileName = "Contrato_Locacao_" + dados.locatario.nome.split(" ")[0] + "_" + new Date().toISOString().split("T")[0] + ".docx";
        
        // Usar saveAs do FileSaver.js ou fallback
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