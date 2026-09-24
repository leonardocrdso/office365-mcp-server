import { describe, expect, it } from "bun:test";
import type { GraphMessageAttachment } from "../types/graph.js";
import { attachmentFileName, selectAttachments } from "./outlook-attachments.js";

describe("selectAttachments", () => {
  const anexoVisivel: GraphMessageAttachment = {
    id: "id-1",
    name: "contrato.pdf",
    contentType: "application/pdf",
    size: 100,
    isInline: false,
    "@odata.type": "#microsoft.graph.fileAttachment",
  };
  const anexoInline: GraphMessageAttachment = {
    id: "id-2",
    name: "logo.png",
    contentType: "image/png",
    size: 50,
    isInline: true,
    "@odata.type": "#microsoft.graph.fileAttachment",
  };
  const anexoLink: GraphMessageAttachment = {
    id: "id-3",
    name: "Planilha compartilhada",
    contentType: "application/json",
    size: 10,
    isInline: false,
    "@odata.type": "#microsoft.graph.referenceAttachment",
  };

  it("sem nome informado, exclui anexos inline e anexos de referencia", () => {
    const selecionados = selectAttachments([anexoVisivel, anexoInline, anexoLink]);
    expect(selecionados).toEqual([anexoVisivel]);
  });

  it("nome exato tem prioridade sobre correspondencia parcial, ignorando maiusculas/minusculas", () => {
    const parcial: GraphMessageAttachment = { ...anexoVisivel, id: "id-4", name: "CONTRATO-2026.pdf" };
    const exato: GraphMessageAttachment = { ...anexoVisivel, id: "id-5", name: "contrato.pdf" };

    const selecionados = selectAttachments([parcial, exato], "contrato.pdf");

    expect(selecionados).toEqual([exato]);
  });

  it("nome inexistente lanca erro listando os nomes disponiveis e mencionando os links quando ha algum", () => {
    expect(() => selectAttachments([anexoVisivel, anexoLink], "inexistente.pdf")).toThrow("'contrato.pdf'");
    expect(() => selectAttachments([anexoVisivel, anexoLink], "inexistente.pdf")).toThrow(
      "Links do OneDrive/SharePoint"
    );
  });

  it("item so com anexos inline ou de link lanca erro", () => {
    expect(() => selectAttachments([anexoInline, anexoLink])).toThrow();
  });
});

describe("attachmentFileName", () => {
  it("acrescenta .eml a anexo do tipo itemAttachment sem extensao", () => {
    const anexoDeEmail: GraphMessageAttachment = {
      id: "id-1",
      name: "Fwd: Proposta comercial",
      contentType: "message/rfc822",
      size: 10,
      isInline: false,
      "@odata.type": "#microsoft.graph.itemAttachment",
    };

    expect(attachmentFileName(anexoDeEmail)).toBe("Fwd: Proposta comercial.eml");
  });

  it("mantem inalterado um anexo de arquivo comum", () => {
    const anexoDeArquivo: GraphMessageAttachment = {
      id: "id-2",
      name: "contrato.pdf",
      contentType: "application/pdf",
      size: 10,
      isInline: false,
      "@odata.type": "#microsoft.graph.fileAttachment",
    };

    expect(attachmentFileName(anexoDeArquivo)).toBe("contrato.pdf");
  });
});
