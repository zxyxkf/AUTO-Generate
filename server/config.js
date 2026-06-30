import fs from 'node:fs';
import path from 'node:path';

export const rootDir = path.resolve(import.meta.dirname, '..');

const exampleSupplierInfo = `Company: Example Supplier Ltd.
Tax ID: EXAMPLE-TAX-ID
Bank: Example Bank
Account: EXAMPLE-ACCOUNT
Address: Example Address
Phone: Example Phone`;

export const themes = {
  example_invoice: {
    label: 'Example Invoice Supplier',
    quotationLabel: 'Example Invoice',
    company: 'Example Supplier Ltd.',
    supplierInfo: exampleSupplierInfo,
    stamp: 'assets/stamps/example-official.png',
    contractStamp: 'assets/stamps/example-contract.png',
    fallbackStamp: '',
    hasSupply: true,
    supplyTemplate: 'excel_templates/example-supply.xlsx',
    contractTemplate: 'word_templates/example-contract.docx',
  },
  example_vat: {
    label: 'Example VAT Supplier',
    quotationLabel: 'Example VAT',
    company: 'Example VAT Supplier Ltd.',
    supplierInfo: exampleSupplierInfo,
    stamp: 'assets/stamps/example-vat-official.png',
    contractStamp: 'assets/stamps/example-vat-contract.png',
    fallbackStamp: '',
    hasSupply: true,
    supplyTemplate: 'excel_templates/example-vat-supply.xlsx',
    contractTemplate: 'word_templates/example-vat-contract.docx',
  },
};

export const documentTypes = {
  quotation: { label: 'Quotation', office: 'excel', template: 'excel_templates/example-quotation.xlsx' },
  supply: { label: 'Supply List', office: 'excel' },
  contract: { label: 'Contract', office: 'word' },
};

export function existingPath(relativePath) {
  if (!relativePath) return '';
  const fullPath = path.join(rootDir, relativePath);
  return fs.existsSync(fullPath) ? fullPath : '';
}
