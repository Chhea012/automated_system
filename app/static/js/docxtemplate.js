// Number to words (enhanced to handle decimals and USD suffix)
function numberToWords(num) {
    try {
        if (!Number.isFinite(num) || num < 0) throw new Error('Invalid number');
        const units = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
        const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
        const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
        const thousands = ['', 'Thousand', 'Million', 'Billion'];

        if (num === 0) return 'Zero US Dollars only';

        // Split into integer and decimal parts
        const integerPart = Math.floor(num);
        const decimalPart = Math.round((num - integerPart) * 100); // Cents

        function convertChunk(n) {
            let chunkText = '';
            if (n >= 100) {
                chunkText += units[Math.floor(n / 100)] + ' Hundred ';
                n %= 100;
            }
            if (n >= 20) {
                chunkText += tens[Math.floor(n / 10)] + ' ';
                n %= 10;
            }
            if (n >= 10 && n < 20) {
                chunkText += teens[n - 10] + ' ';
                n = 0;
            }
            if (n > 0 && n < 10) {
                chunkText += units[n] + ' ';
            }
            return chunkText;
        }

        let result = '';
        let chunkIndex = 0;
        let tempNum = integerPart;
        while (tempNum > 0) {
            const chunk = tempNum % 1000;
            if (chunk > 0) {
                let chunkText = convertChunk(chunk);
                if (chunkText) {
                    result = chunkText + (thousands[chunkIndex] ? thousands[chunkIndex] + ' ' : '') + result;
                }
            }
            tempNum = Math.floor(tempNum / 1000);
            chunkIndex++;
        }

        // Handle cents
        if (decimalPart > 0) {
            result = result.trim() + (result ? ' and ' : '') + convertChunk(decimalPart) + 'Cents ';
        }

        return (result.trim() || 'Zero') + ' US Dollars only';
    } catch (error) {
        console.error('numberToWords error:', error);
        return num.toString();
    }
}

// Date formatting
function formatDate(isoDate) {
    if (!isoDate) return '';
    const date = new Date(isoDate);
    if (isNaN(date)) return isoDate;
    const day = date.getDate();
    const month = date.toLocaleString('en', { month: 'long' });
    const year = date.getFullYear();
    const suffix = (day) => {
        if (day > 3 && day < 21) return 'th';
        switch (day % 10) {
            case 1: return 'st';
            case 2: return 'nd';
            case 3: return 'rd';
            default: return 'th';
        }
    };
    return `${day}${suffix(day)} ${month} ${year}`;
}

// Parse percentage from installment description
function getPercentageFromDesc(desc) {
    const match = desc.match(/\((\d+)%\)/);
    return match ? parseInt(match[1], 10) : 0;
}

function generateDocx(contractData) {
    try {
        console.log('Starting DOCX generation');
        if (!contractData) {
            throw new Error('No contract data available');
        }
        if (!window.docx || !window.saveAs) {
            throw new Error('DOCX or FileSaver libraries not loaded. Ensure docx.min.js and FileSaver.min.js are included.');
        }

        const { contract, totalFee, taxPercentage, taxAmount, netAmount, totalFeeWords } = contractData;
        console.log('Generating DOCX with contract data:', contract);

        // Validate percentages for installments
        let totalPerc = 0;
        const installmentData = contract.payment_installment_desc.map(desc => {
            const perc = getPercentageFromDesc(desc);
            totalPerc += perc;
            const amount = totalFee * (perc / 100);
            const taxAm = amount * (taxPercentage / 100);
            const netAm = amount - taxAm;
            return { desc, amount, taxAm, netAm, perc };
        });
        if (totalPerc !== 100) {
            console.warn('Total percentages do not add up to 100%:', totalPerc);
        }

        const doc = new docx.Document({
            sections: [{
                properties: {
                    page: {
                        margin: { top: 720, right: 720, bottom: 720, left: 720 } // 1 inch margins
                    }
                },
                children: [
                    new docx.Paragraph({
                        text: "The Service Agreement",
                        heading: docx.HeadingLevel.HEADING_1,
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { after: 200 },
                        font: { name: 'Calibri', size: 24 }
                    }),
                    new docx.Paragraph({
                        text: "ON",
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { before: 100, after: 100 },
                        font: { name: 'Calibri', bold: true }
                    }),
                    new docx.Paragraph({
                        text: contract.project_title || '',
                        heading: docx.HeadingLevel.HEADING_2,
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { before: 100, after: 100 },
                        font: { name: 'Calibri', size: 20 }
                    }),
                    new docx.Paragraph({
                        text: `No.: ${contract.contract_number || ''}`,
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { before: 100, after: 100 },
                        font: { name: 'Calibri', bold: true }
                    }),
                    new docx.Paragraph({
                        text: "BETWEEN",
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { before: 100, after: 200 },
                        font: { name: 'Calibri' }
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({ text: "The NGO Forum on Cambodia", bold: true }),
                            new docx.TextRun({ text: ", represented by Mr. Soeung Saroeun, Executive Director.\n" }),
                            new docx.TextRun({ text: "Address: #9-11, Street 476, Sangkat Tuol Tumpoung I, Phnom Penh, Cambodia.\n" }),
                            new docx.TextRun({ text: "hereinafter called the “" }),
                            new docx.TextRun({ text: "Party A", bold: true }),
                            new docx.TextRun({ text: "”" })
                        ],
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { after: 200 },
                        font: { name: 'Calibri' }
                    }),
                    new docx.Paragraph({
                        text: "AND",
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { before: 200, after: 200 },
                        font: { name: 'Calibri' }
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({ text: contract.party_b_signature_name || '', bold: true }),
                            new docx.TextRun({ text: ",\n" }),
                            new docx.TextRun({ text: `Address: ${contract.party_b_address || ''}\n` }),
                            new docx.TextRun({ text: `H/P: ${contract.party_b_phone || ''}, E-mail: ${contract.party_b_email || ''}\n` }),
                            new docx.TextRun({ text: "hereinafter called the “" }),
                            new docx.TextRun({ text: "Party B", bold: true }),
                            new docx.TextRun({ text: "”" })
                        ],
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { after: 200 },
                        font: { name: 'Calibri' }
                    }),
                    new docx.Paragraph({
                        text: "Whereas NGOF is a legal entity registered with the Ministry of Interior (MOI) #304 សជណ dated 07 March 2012.",
                        spacing: { after: 100 },
                        font: { name: 'Calibri' }
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({ text: "Whereas NGOF will engage the services of “" }),
                            new docx.TextRun({ text: "Party B", bold: true }),
                            new docx.TextRun({ text: "” which accept the engagement under the following term and conditions." })
                        ],
                        spacing: { after: 100 },
                        font: { name: 'Calibri' }
                    }),
                    new docx.Paragraph({
                        text: "Both Parties Agreed as follows:",
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { before: 100, after: 200 },
                        font: { name: 'Calibri', bold: true }
                    }),
                    ...[
                        {
                            number: 1,
                            title: 'TERMS OF REFERENCE',
                            content: [
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " shall perform tasks as stated in the attached TOR (annex-1) to " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: ", and deliver each milestone as stipulated in article 4.\n" }),
                                new docx.TextRun({ text: "The work shall be of good quality and well performed with the acceptance by " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: "." })
                            ]
                        },
                        {
                            number: 2,
                            title: 'TERM OF AGREEMENT',
                            content: [
                                new docx.TextRun({ text: `The agreement is effective from ${formatDate(contract.agreement_start_date)} – ${formatDate(contract.agreement_end_date)}.`, bold: true }),
                                new docx.TextRun({ text: "\nThis Agreement is terminated automatically after the due date of the Agreement Term unless otherwise, both Parties agree to extend the Term with a written agreement." })
                            ]
                        },
                        {
                            number: 3,
                            title: 'PROFESSIONAL FEE',
                            content: [
                                new docx.TextRun({ text: `The professional fee is the total amount of USD ${totalFee.toFixed(2)} (${totalFeeWords}) including tax for the whole assignment period.`, bold: true }),
                                new docx.TextRun({ text: `\nTotal Service Fee: USD ${totalFee.toFixed(2)}` }),
                                new docx.TextRun({ text: `\nWithholding Tax ${taxPercentage}%: USD ${taxAmount.toFixed(2)}` }),
                                new docx.TextRun({ text: `\nNet amount: USD ${netAmount.toFixed(2)}` }),
                                new docx.TextRun({ text: "\n" }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " is responsible to issue the Invoice (net amount) and receipt (when receiving the payment) with the total amount as stipulated in each instalment as in the Article 4 after having done the agreed deliverable tasks, for payment request.\n" }),
                                new docx.TextRun({ text: "The payment will be processed after the satisfaction from " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " as of the required deliverable tasks as stated in Article 4.\n" }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " is responsible for all related taxes payable to the government department." })
                            ]
                        },
                        {
                            number: 4,
                            title: 'TERM OF PAYMENT',
                            content: [
                                new docx.TextRun({ text: "The payment will be made based on the following schedules:" })
                            ],
                            table: {
                                rows: [
                                    new docx.TableRow({
                                        children: [
                                            new docx.TableCell({ children: [new docx.Paragraph({ text: "Installment", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                                            new docx.TableCell({ children: [new docx.Paragraph({ text: "Total Amount (USD)", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                                            new docx.TableCell({ children: [new docx.Paragraph({ text: "Deliverable", bold: true, alignment: docx.AlignmentType.CENTER })] }),
                                            new docx.TableCell({ children: [new docx.Paragraph({ text: "Due date", bold: true, alignment: docx.AlignmentType.CENTER })] })
                                        ]
                                    }),
                                    ...installmentData.map(({ desc, amount, taxAm, netAm }) => new docx.TableRow({
                                        children: [
                                            new docx.TableCell({ children: [new docx.Paragraph({ text: desc, alignment: docx.AlignmentType.LEFT })] }),
                                            new docx.TableCell({ children: [
                                                new docx.Paragraph({ text: `· Gross: $${amount.toFixed(2)}` }),
                                                new docx.Paragraph({ text: `· Tax ${taxPercentage}%: $${taxAm.toFixed(2)}` }),
                                                new docx.Paragraph({ text: `· Net pay: $${netAm.toFixed(2)}` })
                                            ] }),
                                            new docx.TableCell({ children: contract.deliverables.map(d => new docx.Paragraph({ text: `· ${d}`, alignment: docx.AlignmentType.LEFT })) }),
                                            new docx.TableCell({ children: [new docx.Paragraph({ text: formatDate(contract.agreement_end_date), alignment: docx.AlignmentType.LEFT })] })
                                        ]
                                    }))
                                ],
                                width: { size: 100, type: docx.WidthType.PERCENTAGE }
                            }
                        },
                        {
                            number: 5,
                            title: 'NO OTHER PERSONS',
                            content: [
                                new docx.TextRun({ text: "No person or entity, which is not a party to this agreement, has any rights to enforce, take any action, or claim it is owed any benefit under this agreement." })
                            ]
                        },
                        {
                            number: 6,
                            title: 'MONITORING and COORDINATION',
                            content: [
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " shall monitor and evaluate the progress of the agreement toward its objective, including the activities implemented.\n" }),
                                new docx.TextRun({ text: `${contract.focal_person_a_name || ''}, ${contract.focal_person_a_position || ''}`, bold: true }),
                                new docx.TextRun({ text: ` (Telephone ${contract.focal_person_a_phone || ''} Email: ${contract.focal_person_a_email || ''}) is the focal contact person of ` }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " and " }),
                                new docx.TextRun({ text: `${contract.party_b_signature_name || ''}`, bold: true }),
                                new docx.TextRun({ text: ` (HP. ${contract.party_b_phone || ''}, E-mail: ${contract.party_b_email || ''}) the focal contact person of ` }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: ".\nThe focal contact person of " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " and " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " will work together for overall coordination including reviewing and meeting discussions during the assignment process." })
                            ]
                        },
                        {
                            number: 7,
                            title: 'CONFIDENTIALITY',
                            content: [
                                new docx.TextRun({ text: `All outputs produced, with the exception of the “${contract.output_description || ''}”, which is a contribution from, and to be claimed as a public document by the main author and co-author in associated, and/or under this agreement, shall be the property of ` }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: ".\nThe " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " agrees to not disclose any confidential information, of which he/she may take cognizance in the performance under this contract, except with the prior written approval of the " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: "." })
                            ]
                        },
                        {
                            number: 8,
                            title: 'ANTI-CORRUPTION and CONFLICT OF INTEREST',
                            content: [
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " shall not participate in any practice that is or could be construed as an illegal or corrupt practice in Cambodia.\nThe " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " is committed to fighting all types of corruption and expects this same commitment from the consultant it reserves the rights and believes based on the declaration of " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " that it is an independent social enterprise firm operating in Cambodia and it does not involve any conflict of interest with other parties that may be affected to the " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: "." })
                            ]
                        },
                        {
                            number: 9,
                            title: 'OBLIGATION TO COMPLY WITH THE NGOF’S POLICIES AND CODE OF CONDUCT',
                            content: [
                                new docx.TextRun({ text: "By signing this agreement, " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " is obligated to comply with and respect all existing policies and code of conduct of " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: ", such as Gender Mainstreaming, Child Protection, Disability policy, Environmental Mainstreaming, etc. and the " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " declared themselves that s/he will perform the assignment in the neutral position, professional manner, and not be involved in any political affiliation." })
                            ]
                        },
                        {
                            number: 10,
                            title: 'ANTI-TERRORISM FINANCING AND FINANCIAL CRIME',
                            content: [
                                new docx.TextRun({ text: "NGOF is determined that all its funds and resources should only be used to further its mission and shall not be subject to illicit use by any third party nor used or abused for any illicit purpose.\n" }),
                                new docx.TextRun({ text: "In order to achieve this objective, NGOF will not knowingly or recklessly provide funds, economic goods, or material support to any entity or individual designated as a “terrorist” by the international community or affiliate domestic governments and will take all reasonable steps to safeguard and protect its assets from such illicit use and to comply with host government laws.\n" }),
                                new docx.TextRun({ text: "NGOF respects its contracts with its donors and puts procedures in place for compliance with these contracts.\n" }),
                                new docx.TextRun({ text: "“Illicit use” refers to terrorist financing, sanctions, money laundering, and export control regulations." })
                            ]
                        },
                        {
                            number: 11,
                            title: 'INSURANCE',
                            content: [
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " is responsible for any health and life insurance of its team members. " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " will not be held responsible for any medical expenses or compensation incurred during or after this contract." })
                            ]
                        },
                        {
                            number: 12,
                            title: 'ASSIGNMENT',
                            content: [
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " shall have the right to assign individuals within its organization to carry out the tasks herein named in the attached Technical Proposal.\n" }),
                                new docx.TextRun({ text: "The " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " shall not assign, or transfer any of its rights or obligations under this agreement hereunder without the prior written consent of " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: ".\n" }),
                                new docx.TextRun({ text: "Any attempt by " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " to assign or transfer any of its rights or obligations without the prior written consent of " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " shall render this agreement subject to immediate termination by " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: "." })
                            ]
                        },
                        {
                            number: 13,
                            title: 'RESOLUTION OF CONFLICTS/DISPUTES',
                            content: [
                                new docx.TextRun({ text: "Conflicts between any of these agreements shall be resolved by the following methods:\n" }),
                                new docx.TextRun({ text: "In the case of a disagreement arising between " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " and the " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " regarding the implementation of any part of, or any other substantive question arising under or relating to this agreement, the parties shall use their best efforts to arrive at an agreeable resolution by mutual consultation.\n" }),
                                new docx.TextRun({ text: "Unresolved issues may, upon the option of either party and written notice to the other party, be referred to for arbitration.\n" }),
                                new docx.TextRun({ text: "Failure by the " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " or " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " to dispute a decision arising from such arbitration in writing within thirty (30) calendar days of receipt of a final decision shall result in such final decision being deemed binding upon either the " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " and/or " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: ".\nAll expenses related to arbitration will be shared equally between both parties." })
                            ]
                        },
                        {
                            number: 14,
                            title: 'TERMINATION',
                            content: [
                                new docx.TextRun({ text: "The " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " or the " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " may, by notice in writing, terminate this agreement under the following conditions:\n" }),
                                new docx.TextRun({ text: "1. " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " may terminate this agreement at any time with a week notice if " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " fails to comply with the terms and conditions of this agreement.\n" }),
                                new docx.TextRun({ text: "2. For gross professional misconduct (as defined in the NGOF Human Resource Policy), " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " may terminate this agreement immediately without prior notice.\n" }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " will notify " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " in a letter that will indicate the reason for termination as well as the effective date of termination.\n" }),
                                new docx.TextRun({ text: "3. " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " may terminate this agreement at any time with a one-week notice if " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " fails to comply with the terms and conditions of this agreement.\n" }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " will notify " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " in a letter that will indicate the reason for termination as well as the effective date of termination.\n" }),
                                new docx.TextRun({ text: "But if " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " intended to terminate this agreement by itself without any appropriate reason or fails of implementing the assignment, " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " has to refund the full amount of fees received to " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: ".\n" }),
                                new docx.TextRun({ text: "4. If for any reason either " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " or the " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " decides to terminate this agreement, " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " shall be paid pro-rata for the work already completed by " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: ".\n" }),
                                new docx.TextRun({ text: "This payment will require the submission of a timesheet that demonstrates work completed as well as the handing over of any deliverables completed or partially completed.\n" }),
                                new docx.TextRun({ text: "In case " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " has received payment for services under the agreement which have not yet been performed; the appropriate portion of these fees would be refunded by " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: " to " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: "." })
                            ]
                        },
                        {
                            number: 15,
                            title: 'MODIFICATION OR AMENDMENT',
                            content: [
                                new docx.TextRun({ text: "No modification or amendment of this agreement shall be valid unless in writing and signed by an authorized person of " }),
                                new docx.TextRun({ text: "Party A", bold: true }),
                                new docx.TextRun({ text: " and " }),
                                new docx.TextRun({ text: "Party B", bold: true }),
                                new docx.TextRun({ text: "." })
                            ]
                        },
                        {
                            number: 16,
                            title: 'CONTROLLING OF LAW',
                            content: [
                                new docx.TextRun({ text: "This agreement shall be governed and construed following the law of the Kingdom of Cambodia.\n" }),
                                new docx.TextRun({ text: "The Simultaneous Interpretation Agreement is prepared in two original copies." })
                            ]
                        }
                    ].flatMap(article => [
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({ text: `ARTICLE ${article.number}`, bold: true, underline: { type: docx.UnderlineType.SINGLE } }),
                                new docx.TextRun({ text: `: ${article.title}`, bold: true })
                            ],
                            spacing: { after: 100 },
                            font: { name: 'Calibri', size: 14 }
                        }),
                        new docx.Paragraph({
                            children: article.content,
                            spacing: { after: 100 },
                            font: { name: 'Calibri' }
                        }),
                        ...(article.table ? [new docx.Table({
                            ...article.table,
                            width: { size: 100, type: docx.WidthType.PERCENTAGE }
                        })] : []),
                        ...contract.articles.filter(a => parseInt(a.article_number, 10) === article.number).map(a => new docx.Paragraph({
                            text: a.custom_sentence || '',
                            spacing: { after: 100 },
                            font: { name: 'Calibri' }
                        }))
                    ]),
                    new docx.Paragraph({
                        text: `Date: ${formatDate(contract.agreement_start_date)}`,
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { before: 200, after: 200 },
                        font: { name: 'Calibri', bold: true }
                    }),
                    new docx.Table({
                        rows: [
                            new docx.TableRow({
                                children: [
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({ text: "For “Party A”", alignment: docx.AlignmentType.CENTER, font: { name: 'Calibri', bold: true } }),
                                            new docx.Paragraph({ text: "_________________", alignment: docx.AlignmentType.CENTER, spacing: { before: 2000 }, font: { name: 'Calibri', bold: true } }),
                                            new docx.Paragraph({ text: "Mr. SOEUNG Saroeun", alignment: docx.AlignmentType.CENTER, font: { name: 'Calibri', bold: true } }),
                                            new docx.Paragraph({ text: "Executive Director", alignment: docx.AlignmentType.CENTER, font: { name: 'Calibri', bold: true } })
                                        ],
                                        width: { size: 50, type: docx.WidthType.PERCENTAGE }
                                    }),
                                    new docx.TableCell({
                                        children: [
                                            new docx.Paragraph({ text: "For “Party B”", alignment: docx.AlignmentType.CENTER, font: { name: 'Calibri', bold: true } }),
                                            new docx.Paragraph({ text: "____________________", alignment: docx.AlignmentType.CENTER, spacing: { before: 2000 }, font: { name: 'Calibri', bold: true } }),
                                            new docx.Paragraph({ text: contract.party_b_signature_name || '', alignment: docx.AlignmentType.CENTER, font: { name: 'Calibri', bold: true } }),
                                            new docx.Paragraph({ text: contract.party_b_position || '', alignment: docx.AlignmentType.CENTER, font: { name: 'Calibri', bold: true } })
                                        ],
                                        width: { size: 50, type: docx.WidthType.PERCENTAGE }
                                    })
                                ]
                            })
                        ],
                        width: { size: 100, type: docx.WidthType.PERCENTAGE }
                    })
                ]
            }]
        });

        console.log('DOCX document created, attempting to generate Blob');
        docx.Packer.toBlob(doc).then(blob => {
            console.log('Blob created successfully, size:', blob.size);
            saveAs(blob, `Service_Agreement_${contract.contract_number || 'unknown'}.docx`);
            console.log('DOCX file download initiated');
        }).catch(err => {
            console.error('Error generating Blob:', err);
            throw new Error('Failed to generate DOCX Blob: ' + err.message);
        });
    } catch (error) {
        console.error('generateDocx error:', error);
        throw error;
    }
}