import React, { useState } from "react";
import { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat, PageBreak } from "docx";
import { FileText, User, Briefcase, Download, AlertCircle, Building2 } from "lucide-react";

const COMPANY = {
  name: "Skyfire Systems Inc.",
  state: "Delaware",
  cfo: "Todd Parker, CFO",
  addr1: "77 Geary St.",
  addr2: "5th Floor",
  cityState: "San Francisco, CA 94108",
  email: "Todd@Skyfire.xyz",
};

const formatDate = (iso) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  return `${months[parseInt(m,10)-1]} ${parseInt(d,10)}, ${y}`;
};

const numberingConfig = {
  config: [
    { reference: "topnum", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "sublist", levels: [{ level: 0, format: LevelFormat.LOWER_LETTER, text: "(%1)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } }] },
  ],
};

const baseStyles = {
  default: { document: { run: { font: "Calibri", size: 22 } } },
};

const plain = (text, opts = {}) => new Paragraph({ spacing: { before: 100, after: 100 }, children: [new TextRun({ text, ...opts })] });
const center = (text, opts = {}) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 }, children: [new TextRun({ text, ...opts })] });
const blank = () => new Paragraph({ children: [new TextRun("")] });

const numberedSection = (heading, body) => new Paragraph({
  numbering: { reference: "topnum", level: 0 },
  spacing: { before: 160, after: 120 },
  children: [new TextRun({ text: heading, bold: true }), new TextRun({ text: "  " }), new TextRun({ text: body })],
});

const subSection = (heading, body) => new Paragraph({
  numbering: { reference: "sublist", level: 0 },
  spacing: { before: 120, after: 120 },
  children: [new TextRun({ text: heading, bold: true }), new TextRun({ text: "  " }), new TextRun({ text: body })],
});

// ============================================================
// CONSULTANT DOC — verbatim from template
// ============================================================
function buildConsultantDoc(data) {
  const eff = formatDate(data.effectiveDate) || "[DATE]";
  const name = data.name || "[CONSULTANT NAME]";
  const addr = data.address || "";
  const email = data.email || "";
  const phone = data.phone || "";

  const c = [];

  c.push(center(COMPANY.name, { bold: true, size: 28 }));
  c.push(blank());
  c.push(center("CONFIDENTIAL INFORMATION AND", { bold: true, size: 26 }));
  c.push(center("INVENTION ASSIGNMENT AGREEMENT", { bold: true, size: 26 }));
  c.push(blank());

  c.push(new Paragraph({ children: [
    new TextRun({ text: "Consultant Name: ", bold: true }),
    new TextRun({ text: `${name} (\u201CConsultant\u201D)` }),
  ]}));
  c.push(new Paragraph({ children: [
    new TextRun({ text: "Effective Date: ", bold: true }),
    new TextRun({ text: eff }),
  ]}));
  c.push(blank());

  c.push(plain(`As a condition of becoming retained (or Consultant\u2019s consulting relationship being continued) by ${COMPANY.name}, a ${COMPANY.state} corporation, or any of its current or future subsidiaries, affiliates, successors or assigns (collectively, the \u201CCompany\u201D), and in consideration of Consultant\u2019s consulting relationship with the Company and receipt of the compensation now and hereafter paid by the Company, the receipt of Confidential Information (as defined below) while associated with the Company, and other good and valuable consideration, the receipt and sufficiency of which are hereby acknowledged, Consultant agrees to the following:`));

  c.push(numberedSection("Relationship.", `This Confidential Information and Invention Assignment Agreement (this \u201CAgreement\u201D) will apply to Consultant\u2019s consulting relationship with the Company. If that relationship ends and the Company, within a year thereafter, either employs Consultant or re-engages Consultant as a consultant, this Agreement will also apply to such later employment or consulting relationship, unless the parties hereto otherwise agree in writing. Any such employment or consulting relationship between the parties hereto, whether commenced prior to, upon or after the date of this Agreement, is referred to herein as the \u201CRelationship.\u201D`));

  c.push(numberedSection("Consulting Agreement.", `Consultant has entered into an agreement with the Company on or about the date hereof to provide various services to the Company (the \u201CConsulting Agreement\u201D). The services to be rendered by Consultant under the Consulting Agreement are referred to herein as the \u201CServices\u201D and this Agreement is intended to supplement and form an integral part of the Consulting Agreement.`));

  c.push(numberedSection("Confidential Information.", ""));
  c.push(subSection("Protection of Information.", `Consultant understands that during the Relationship, the Company intends to provide Consultant with information, including Confidential Information (as defined below), without which Consultant would not be able to perform Consultant\u2019s duties to the Company. Consultant agrees, at all times during the term of the Relationship and thereafter, to hold in strictest confidence, and not to use, except for the benefit of the Company to the extent necessary to perform the Services, and not to disclose to any person, firm, corporation or other entity, without written authorization from the Company in each instance, any Confidential Information that Consultant obtains from the Company or otherwise obtains, accesses or creates in connection with, or as a result of, the Services prior to or during the term of the Relationship, whether or not during working hours, until such Confidential Information becomes publicly and widely known and made generally available through no wrongful act of Consultant or of others who were under confidentiality obligations as to the item or items involved. Consultant further agrees not to make copies of such Confidential Information except as authorized by the Company.`));
  c.push(subSection("Confidential Information.", `Consultant understands that \u201CConfidential Information\u201D means information and physical material not generally known or available outside the Company and information and physical material entrusted to the Company in confidence by third parties. Confidential Information includes, without limitation: (i) Company Inventions (as defined below); and (ii) technical data, trade secrets, know-how, research, product or service ideas or plans, software codes and designs, algorithms, developments, inventions, patent applications, laboratory notebooks, processes, formulas, techniques, biological materials, mask works, engineering designs and drawings, hardware configuration information, agreements with third parties, lists of, or information relating to, employees and consultants of the Company (including, but not limited to, the names, contact information, jobs, compensation, and expertise of such employees and consultants), lists of, or information relating to, suppliers and customers (including, but not limited to, customers of the Company on whom Consultant called or with whom Consultant became acquainted during the Relationship), price lists, pricing methodologies, cost data, market share data, marketing plans, licenses, contract information, business plans, financial forecasts, historical financial data, budgets or other business information disclosed to Consultant by the Company either directly or indirectly, whether in writing, electronically, orally, or by observation.`));
  c.push(subSection("Third Party Information.", `Consultant\u2019s agreements in this Section 3 are intended to be for the benefit of the Company and any third party that has entrusted information or physical material to the Company in confidence. Consultant further agrees that, during the term of the Relationship and thereafter, Consultant will not improperly use or disclose to the Company any confidential, proprietary or secret information of Consultant\u2019s former clients or any other person, and Consultant agrees not to bring any such information onto the Company\u2019s property or place of business.`));
  c.push(subSection("Other Rights.", `This Agreement is intended to supplement, and not to supersede, any rights the Company may have in law or equity with respect to the protection of trade secrets or confidential or proprietary information.`));
  c.push(subSection("U.S. Defend Trade Secrets Act.", `Notwithstanding the foregoing, the U.S. Defend Trade Secrets Act of 2016 (\u201CDTSA\u201D) provides that an individual shall not be held criminally or civilly liable under any federal or state trade secret law for the disclosure of a trade secret that is made (i) in confidence to a federal, state, or local government official, either directly or indirectly, or to an attorney; and (ii) solely for the purpose of reporting or investigating a suspected violation of law; or (iii) in a complaint or other document filed in a lawsuit or other proceeding, if such filing is made under seal. In addition, DTSA provides that an individual who files a lawsuit for retaliation by an employer for reporting a suspected violation of law may disclose the trade secret to the attorney of the individual and use the trade secret information in the court proceeding, if the individual (A) files any document containing the trade secret under seal; and (B) does not disclose the trade secret, except pursuant to court order.`));

  c.push(numberedSection("Ownership of Inventions.", ""));
  c.push(subSection("Inventions Retained and Licensed.", `Consultant has attached hereto, as Exhibit A, a complete list describing with particularity all Inventions (as defined below) that, as of the Effective Date: (i) Consultant made, and/or (ii) belong solely to Consultant or belong to Consultant jointly with others or in which Consultant has an interest, and that relate in any way to any of the Company\u2019s actual or proposed businesses, products, services, or research and development, and which are not assigned to the Company hereunder; or, if no such list is attached, Consultant represents that there are no such Inventions at the time of signing this Agreement, and to the extent such Inventions do exist and are not listed on Exhibit A, Consultant hereby forever waives any and all rights or claims of ownership to such Inventions. Consultant understands that Consultant\u2019s listing of any Inventions on Exhibit A does not constitute an acknowledgement by the Company of the existence or extent of such Inventions, nor of Consultant\u2019s ownership of such Inventions. Consultant further understands that Consultant must receive the formal approval of the Company before commencing Consultant\u2019s Relationship with the Company.`));
  c.push(subSection("Use or Incorporation of Inventions.", `If in the course of the Relationship, Consultant uses or incorporates into a product, service, process or machine any Invention not covered by Section 4(d) of this Agreement in which Consultant has an interest, Consultant will promptly so inform the Company in writing. Whether or not Consultant gives such notice, Consultant hereby irrevocably grants to the Company a nonexclusive, fully paid-up, royalty-free, assumable, perpetual, worldwide license, with right to transfer and to sublicense, to practice and exploit such Invention and to make, have made, copy, modify, make derivative works of, use, sell, import, and otherwise distribute such Invention under all applicable intellectual property laws without restriction of any kind.`));
  c.push(subSection("Inventions.", `Consultant understands that \u201CInventions\u201D means discoveries, developments, concepts, designs, ideas, know how, improvements, inventions, trade secrets and/or original works of authorship, whether or not patentable, copyrightable or otherwise legally protectable. Consultant understands this includes, but is not limited to, any new product, machine, article of manufacture, biological material, method, procedure, process, technique, use, equipment, device, apparatus, system, compound, formulation, composition of matter, design or configuration of any kind, or any improvement thereon. Consultant understands that \u201CCompany Inventions\u201D means any and all Inventions that Consultant or Consultant\u2019s personnel may solely or jointly author, discover, develop, conceive, or reduce to practice in connection with, or as a result of, the Services performed for the Company during or before the term of this Agreement, except as otherwise provided in Section 4(g) below. Without limiting the foregoing, Company Inventions include all deliverables specified in any Project Assignment under the Consulting Agreement.`));
  c.push(subSection("Assignment of Company Inventions.", `Consultant hereby assigns to the Company, or its designee, and Consultant agrees that Consultant will promptly make full written disclosure to the Company, of and to hold in trust for the sole right and benefit of the Company, all Consultant\u2019s right, title and interest throughout the world in and to any and all Company Inventions and all patent, copyright, trademark, trade secret and other intellectual property rights therein. Consultant hereby waives and irrevocably quitclaims to the Company or its designee any and all claims, of any nature whatsoever, that Consultant now has or may hereafter have for infringement of any and all Company Inventions. Any assignment of Company Inventions includes all rights of attribution, paternity, integrity, modification, disclosure and withdrawal, and any other rights throughout the world that may be known as or referred to as \u201Cmoral rights,\u201D \u201Cartist\u2019s rights,\u201D \u201Cdroit moral,\u201D or the like (collectively, \u201CMoral Rights\u201D). To the extent that Moral Rights cannot be assigned under applicable law, Consultant hereby waives and agrees not to enforce any and all Moral Rights, including, without limitation, any limitation on subsequent modification, to the extent permitted under applicable law.`));
  c.push(subSection("Maintenance of Records.", `Consultant agrees to keep and maintain adequate and current written records of all Company Inventions made or conceived by Consultant or Consultant\u2019s personnel (solely or jointly with others) during the term of the Relationship. The records may be in the form of notes, sketches, drawings, flow charts, electronic data or recordings, laboratory notebooks, or any other format. The records will be available to and remain the sole property of the Company at all times. Consultant agrees not to remove such records from the Company\u2019s place of business except as expressly permitted by Company policy which may, from time to time, be revised at the sole election of the Company for the purpose of furthering the Company\u2019s business. Consultant agrees to deliver all such records (including any copies thereof) to the Company at the time of termination of the Relationship as provided for in Section 5 and Section 6.`));
  c.push(subSection("Patent and Copyright Rights.", `Consultant agrees to assist the Company, or its designee, at its expense, in every proper way to secure the Company\u2019s, or its designee\u2019s, rights in the Company Inventions and any copyrights, patents, trademarks, mask work rights, Moral Rights, or other intellectual property rights relating thereto in any and all countries, including the disclosure to the Company or its designee of all pertinent information and data with respect thereto, the execution of all applications, specifications, oaths, assignments, recordations, and all other instruments which the Company or its designee shall deem necessary in order to apply for, obtain, maintain and transfer such rights, or if not transferable, waive and agree never to assert such rights, and in order to assign and convey to the Company or its designee, and any successors, assigns and nominees the sole and exclusive right, title and interest in and to such Company Inventions, and any copyrights, patents, mask work rights or other intellectual property rights relating thereto. Consultant further agrees that Consultant\u2019s obligation to execute or cause to be executed, when it is in Consultant\u2019s power to do so, any such instrument or papers shall continue during and at all times after the end of the Relationship and until the expiration of the last such intellectual property right to expire in any country of the world. Consultant hereby irrevocably designates and appoints the Company and its duly authorized officers and agents as Consultant\u2019s agent and attorney-in-fact, to act for and in Consultant\u2019s behalf and stead to execute and file any such instruments and papers and to do all other lawfully permitted acts to further the application for, prosecution, issuance, maintenance or transfer of letters patent, copyright, mask work and other registrations related to such Company Inventions. This power of attorney is coupled with an interest and shall not be affected by Consultant\u2019s subsequent incapacity.`));
  c.push(subSection("Exception to Assignments.", `Subject to the requirements of applicable law, if any, Consultant understands that the Company Inventions will not include, and the provisions of this Agreement requiring assignment of inventions to the Company do not apply to, any invention which qualifies fully for exclusion under the provisions of applicable law, if any, attached hereto as Exhibit B. In order to assist in the determination of which inventions qualify for such exclusion, Consultant will advise the Company promptly in writing, during and for a period of twelve (12) months immediately following the termination of the Relationship, of all Inventions solely or jointly conceived or developed or reduced to practice by Consultant or Consultant\u2019s personnel in connection with, or as a result of, the Services performed for the Company during the period of the Relationship.`));

  c.push(numberedSection("Company Property; Returning Company Documents.", `Consultant acknowledges and agrees that Consultant has no expectation of privacy with respect to the Company\u2019s telecommunications, networking or information processing systems (including, without limitation, files, e-mail messages, and voice messages) and that Consultant\u2019s activity and any files or messages on or using any of those systems may be monitored or reviewed at any time without notice. Consultant further agrees that any property situated on the Company\u2019s premises and owned by the Company, including disks and other storage media, filing cabinets or other work areas, is subject to inspection by Company personnel at any time with or without notice. Consultant agrees that, at the time of termination of the Relationship, Consultant will deliver to the Company (and will not keep in Consultant\u2019s possession, recreate or deliver to anyone else) any and all devices, records, data, notes, reports, proposals, lists, correspondence, specifications, drawings, blueprints, sketches, laboratory notebooks, materials, flow charts, equipment, other documents or property, or reproductions of any of the aforementioned items developed by Consultant or Consultant\u2019s personnel pursuant to the Relationship or otherwise belonging to the Company, its successors or assigns.`));

  c.push(numberedSection("Termination Certification.", `In the event of the termination of the Relationship, Consultant agrees to sign and deliver the \u201CTermination Certification\u201D attached hereto as Exhibit C; however, Consultant\u2019s failure to sign and deliver the Termination Certification shall in no way diminish Consultant\u2019s continuing obligations under this Agreement.`));

  c.push(numberedSection("Notice to Third Parties.", `Consultant agrees that during the periods of time during which Consultant is restricted in taking certain actions by the terms of Section 8 of this Agreement (the \u201CRestriction Period\u201D), Consultant shall inform any entity or person with whom Consultant may seek to enter into a business relationship (whether as an owner, employee, independent contractor or otherwise) of Consultant\u2019s contractual obligations under this Agreement. Consultant also understands and agrees that the Company may, with or without prior notice to Consultant and during or after the term of the Relationship, notify third parties of Consultant\u2019s agreements and obligations under this Agreement. Consultant further agrees that, upon written request by the Company, Consultant will respond to the Company in writing regarding the status of Consultant\u2019s engagement or proposed engagement with any party during the Restriction Period.`));

  c.push(numberedSection("Solicitation of Employees, Consultants and Other Parties.", `As described above, Consultant acknowledges and agrees that the Company\u2019s Confidential Information includes information relating to the Company\u2019s employees, consultants, customers and others, and that Consultant will not use or disclose such Confidential Information except as authorized by the Company. Consultant further agrees as follows:`));
  c.push(subSection("Employees, Consultants.", `Consultant agrees that during the term of the Relationship, and for a period of twelve (12) months immediately following the termination of the Relationship for any reason, whether with or without cause, Consultant shall not, directly or indirectly, solicit any of the Company\u2019s employees or consultants to terminate their relationship with the Company, or attempt to solicit or take away employees or consultants of the Company, either for Consultant or for any other person or entity.`));
  c.push(subSection("Other Parties.", `Consultant agrees that during the term of the Relationship, Consultant will not negatively influence any of the Company\u2019s clients, licensors, licensees or customers from purchasing Company products or services or solicit or influence or attempt to influence any client, licensor, licensee, customer or other person either directly or indirectly, to direct any purchase of products and/or services to any person, firm, corporation, institution or other entity in competition with the business of the Company.`));

  c.push(numberedSection("No Change to Duration of Relationship.", `Consultant understands and acknowledges that this Agreement does not alter, amend or expand upon any rights Consultant may have to continue in the consulting relationship with, or in the duration of Consultant\u2019s consulting relationship with, the Company under any existing agreements between the Company and Consultant, including without limitation the Consulting Agreement, or under applicable law.`));

  c.push(numberedSection("Representations and Covenants.", ""));
  c.push(subSection("Facilitation of Agreement.", `Consultant agrees to execute promptly, both during and after the end of the Relationship, any proper oath, and to verify any proper document, required to carry out the terms of this Agreement, upon the Company\u2019s written request to do so.`));
  c.push(subSection("No Conflicts.", `Consultant represents that Consultant\u2019s performance of all the terms of this Agreement does not and will not breach any agreement Consultant has entered into, or will enter into, with any third party, including without limitation any agreement to keep in confidence proprietary information or materials acquired by Consultant in confidence or in trust prior to or during the Relationship. Consultant will not disclose to the Company or use any inventions, confidential or non-public proprietary information or material belonging to any previous client, employer or any other party. Consultant will not induce the Company to use any inventions, confidential or non-public proprietary information, or material belonging to any previous client, employer or any other party. Consultant acknowledges and agrees that Consultant has listed on Exhibit D all agreements (e.g., non-competition agreements, non-solicitation of customers agreements, non-solicitation of employees agreements, confidentiality agreements, inventions agreements, etc.), if any, with a current or former client, employer, or any other person or entity, that may restrict Consultant\u2019s ability to perform services for the Company or Consultant\u2019s ability to recruit or engage customers or service providers on behalf of the Company, or otherwise relate to or restrict Consultant\u2019s ability to perform Consultant\u2019s duties for the Company or any obligation Consultant may have to the Company. Consultant agrees not to enter into any written or oral agreement that conflicts with the provisions of this Agreement. Consultant further represents that Consultant does not presently perform or intend to perform, during the term of the Consulting Agreement, consulting or other services for, and Consultant is not presently employed by and has no intention of being employed by, companies whose businesses or proposed businesses in any way involve products or services that would be competitive with the Company\u2019s products or services, or those products or services proposed or in development by the Company during the term of the Consulting Agreement (except for those companies, if any, listed on Exhibit D attached hereto). If, however, Consultant decides to do so, Consultant agrees that, in advance of accepting such employment or agreeing to perform such services, Consultant will promptly notify the Company in writing, specifying the organization to which Consultant proposes to render services, and provide information sufficient to allow the Company to determine if such work would conflict with the interests of the Company.`));
  c.push(subSection("Voluntary Execution.", `Consultant certifies and acknowledges that Consultant has carefully read all of the provisions of this Agreement, that Consultant understands and has voluntarily accepted such provisions, and that Consultant will fully and faithfully comply with such provisions.`));

  c.push(numberedSection("Miscellaneous.", ""));
  c.push(subSection("Governing Law.", `The validity, interpretation, construction and performance of this Agreement, and all acts and transactions pursuant hereto and the rights and obligations of the parties hereto shall be governed, construed and interpreted in accordance with the laws of the State of ${COMPANY.state} USA, without giving effect to the principles of conflict of laws. The parties agree that any and all disputes arising out of the terms of this Agreement or their interpretation shall be subject to binding arbitration held in San Francisco, California USA before a single arbitrator of the American Arbitration Association in accordance with its Rules of Arbitration and pursuant to ${COMPANY.state} law. The parties agree that the prevailing party in any arbitration shall be entitled to injunctive relief in any court of competent jurisdiction to enforce the arbitration award.`));
  c.push(subSection("Entire Agreement.", `Except as described in Section 2, this Agreement sets forth the entire agreement and understanding between the Company and Consultant relating to its subject matter and merges all prior discussions between the parties to this Agreement. No amendment to this Agreement will be effective unless in writing signed by both parties to this Agreement. The Company shall not be deemed hereby to have waived any rights or remedies it may have in law or equity, nor to have given any authorizations or waived any of its rights under this Agreement, unless, and only to the extent, it does so by a specific writing signed by a duly authorized officer of the Company. Any subsequent change or changes in Consultant\u2019s duties, obligations, rights or compensation will not affect the validity or scope of this Agreement.`));
  c.push(subSection("Successors and Assigns.", `This Agreement will be binding upon Consultant\u2019s successors and assigns, and will be for the benefit of the Company, its successors, and its assigns.`));
  c.push(subSection("Notices.", `Any notice, demand or request required or permitted to be given under this Agreement shall be in writing and shall be deemed sufficient when delivered personally or by overnight courier or sent by email, or 48 hours after being deposited in the U.S. mail as certified or registered mail with postage prepaid, addressed to the party to be notified at such party\u2019s address as set forth on the signature page, as subsequently modified by written notice, or if no address is specified on the signature page, at the most recent address set forth in the Company\u2019s books and records.`));
  c.push(subSection("Severability.", `If one or more of the provisions in this Agreement are deemed void or unenforceable to any extent in any context, such provisions shall nevertheless be enforced to the fullest extent allowed by law in that and other contexts, and the validity and force of the remainder of this Agreement shall not be affected. The Company and Consultant have attempted to limit Consultant\u2019s right to use, maintain and disclose the Company\u2019s Confidential Information, and to limit Consultant\u2019s right to solicit employees and customers only to the extent necessary to protect the Company from unfair competition. Should a court of competent jurisdiction determine that the scope of the covenants contained in Section 8 exceeds the maximum restrictiveness such court deems reasonable and enforceable, the parties intend that the court should reform, modify and enforce the provision to such narrower scope as it determines to be reasonable and enforceable under the circumstances existing at that time. In the event that any court or government agency of competent jurisdiction determines that, notwithstanding the terms of the Consulting Agreement specifying Consultant\u2019s Relationship with the Company as that of an independent contractor, Consultant\u2019s provision of services to the Company is not as an independent contractor but instead as an employee under the applicable laws, then solely to the extent that such determination is applicable, references in this Agreement to the Relationship between Consultant and the Company shall be interpreted to include an employment relationship, and this Agreement shall not be invalid and unenforceable but shall be read to the fullest extent as may be valid and enforceable under the applicable laws to carry out the intent and purpose of the Agreement.`));
  c.push(subSection("Remedies.", `Consultant acknowledges and agrees that violation of this Agreement by Consultant may cause the Company irreparable harm, and therefore Consultant agrees that the Company will be entitled to seek extraordinary relief in court, including, but not limited to, temporary restraining orders, preliminary injunctions and permanent injunctions without the necessity of posting a bond or other security (or, where such a bond or security is required, Consultant agrees that a $1,000 bond will be adequate), in addition to and without prejudice to any other rights or remedies that the Company may have for a breach of this Agreement.`));
  c.push(subSection("Advice of Counsel.", `CONSULTANT ACKNOWLEDGES THAT, IN EXECUTING THIS AGREEMENT, CONSULTANT HAS HAD THE OPPORTUNITY TO SEEK THE ADVICE OF INDEPENDENT LEGAL COUNSEL, AND CONSULTANT HAS READ AND UNDERSTANDS ALL OF THE TERMS AND PROVISIONS OF THIS AGREEMENT. THIS AGREEMENT SHALL NOT BE CONSTRUED AGAINST ANY PARTY BY REASON OF THE DRAFTING OR PREPARATION HEREOF.`));
  c.push(subSection("Electronic Delivery.", `The Company may, in its sole discretion, decide to deliver any documents related to this Agreement or any notices required by applicable law by email or any other electronic means. Consultant hereby consents to (i) conduct business electronically (ii) receive such documents and notices by such electronic delivery and (iii) sign documents electronically and agrees to participate through an on-line or electronic system established and maintained by the Company or a third party designated by the Company.`));
  c.push(subSection("Counterparts.", `This Agreement may be executed in any number of counterparts, each of which when so executed and delivered shall be deemed an original, and all of which together shall constitute one and the same agreement. Execution of a facsimile copy will have the same force and effect as execution of an original, and a facsimile signature will be deemed an original and valid signature.`));

  c.push(blank());
  c.push(center("[Signature Page Follows]", { italics: true }));
  c.push(new Paragraph({ children: [new PageBreak()] }));

  // Signature page
  c.push(plain("The parties have executed this Agreement on the respective dates set forth below, to be effective as of the Effective Date first above written."));
  c.push(blank());
  c.push(plain("THE COMPANY:", { bold: true }));
  c.push(plain(COMPANY.name, { bold: true }));
  c.push(blank());
  c.push(plain("By: ___________________________"));
  c.push(plain(`        ${COMPANY.cfo}`));
  c.push(plain(`        ${COMPANY.name}`));
  c.push(plain(COMPANY.addr1));
  c.push(plain(COMPANY.addr2));
  c.push(plain(COMPANY.cityState));
  c.push(plain(`Email: ${COMPANY.email}`));
  c.push(blank());
  c.push(new Paragraph({ children: [
    new TextRun({ text: "CONSULTANT: ", bold: true }),
    new TextRun({ text: name }),
  ]}));
  c.push(blank());
  c.push(plain("By: ____________________________"));
  c.push(plain(`Name: ${name}`));
  c.push(plain(`Address: ${addr}`));
  c.push(plain(`Email: ${email}`));
  c.push(plain(`Phone: ${phone}`));

  // EXHIBIT A
  c.push(new Paragraph({ children: [new PageBreak()] }));
  c.push(center("EXHIBIT A", { bold: true, size: 28 }));
  c.push(blank());
  c.push(center("LIST OF PRIOR INVENTIONS", { bold: true }));
  c.push(center("AND ORIGINAL WORKS OF AUTHORSHIP", { bold: true }));
  c.push(center("EXCLUDED UNDER SECTION 4(a)", { bold: true }));
  c.push(blank());
  c.push(plain(`The following is a list of all Inventions that, as of the Effective Date: (A) Consultant made, and/or (B) belong solely to Consultant or belong to Consultant jointly with others or in which Consultant has an interest, and that relate in any way to any of the Company\u2019s actual or proposed businesses, products, services, or research and development, and which are not assigned to the Company:`));
  c.push(blank());
  c.push(plain("Prior Inventions:", { bold: true }));
  const prior = (data.priorInventions || "").trim();
  if (prior) {
    prior.split(/\r?\n/).forEach(line => c.push(plain(line)));
  } else {
    c.push(plain("None."));
  }
  c.push(blank());
  c.push(plain(`Except as indicated above on this Exhibit A, Consultant has no inventions, improvements or original works to disclose pursuant to Section 4(a) of this Agreement.`));
  c.push(blank());
  c.push(new Paragraph({ children: [new TextRun({ text: "CONSULTANT: ", bold: true }), new TextRun({ text: name })] }));
  c.push(blank());
  c.push(plain("By: ___________________________"));
  c.push(plain(`Name: ${name}`));
  c.push(plain(`Address: ${addr}`));
  c.push(plain(`Email: ${email}`));
  c.push(plain(`Phone: ${phone}`));

  // EXHIBIT B
  c.push(new Paragraph({ children: [new PageBreak()] }));
  c.push(center("EXHIBIT B", { bold: true, size: 28 }));
  c.push(blank());
  c.push(plain("None."));

  // EXHIBIT C
  c.push(new Paragraph({ children: [new PageBreak()] }));
  c.push(center("EXHIBIT C", { bold: true, size: 28 }));
  c.push(blank());
  c.push(center("TERMINATION CERTIFICATION", { bold: true }));
  c.push(blank());
  c.push(plain(`This is to certify that Consultant does not have in Consultant\u2019s possession, nor has Consultant failed to return, any devices, records, data, notes, reports, proposals, lists, correspondence, specifications, drawings, blueprints, sketches, laboratory notebooks, flow charts, materials, equipment, other documents or property, or copies or reproductions of any aforementioned items belonging to ${COMPANY.name}, a ${COMPANY.state} corporation, its subsidiaries, affiliates, successors or assigns (collectively, the \u201CCompany\u201D).`));
  c.push(plain(`Consultant further certifies that Consultant has complied with all the terms of the Company\u2019s Confidential Information and Invention Assignment Agreement signed by Consultant, including the reporting of any Inventions (as defined therein), conceived or made by Consultant or Consultant\u2019s personnel (solely or jointly with others) covered by that agreement, and Consultant acknowledges Consultant\u2019s continuing obligations under that agreement.`));
  c.push(plain(`Consultant further agrees that, in compliance with the Confidential Information and Invention Assignment Agreement, Consultant will preserve as confidential all trade secrets, confidential knowledge, data or other proprietary information relating to products, processes, know-how, designs, formulas, developmental or experimental work, computer programs, data bases, other original works of authorship, customer lists, business plans, financial information or other subject matter pertaining to any business of the Company or any of its employees, clients, consultants or licensees.`));
  c.push(plain(`Consultant further agrees that for twelve (12) months immediately following the termination of Consultant\u2019s Relationship with the Company, Consultant shall not either directly or indirectly solicit any of the Company\u2019s employees or consultants to terminate their relationship with the Company, or attempt to solicit or take away employees or consultants of the Company, either for Consultant or for any other person or entity.`));
  c.push(plain(`Further, Consultant agrees that Consultant shall not use any Confidential Information of the Company to negatively influence any of the Company\u2019s clients or customers from purchasing Company products or services or to solicit or influence or attempt to influence any client, customer or other person either directly or indirectly, to direct any purchase of products and/or services to any person, firm, corporation, institution or other entity in competition with the business of the Company.`));
  c.push(blank());
  c.push(new Paragraph({ children: [new TextRun({ text: "Consultant: ", bold: true }), new TextRun({ text: `_____________________` })] }));
  c.push(blank());
  c.push(plain("Signature: _______________________________     Date: _____________"));

  // EXHIBIT D
  c.push(new Paragraph({ children: [new PageBreak()] }));
  c.push(center("EXHIBIT D", { bold: true, size: 28 }));
  c.push(blank());
  c.push(center("LIST OF Agreements and Companies", { bold: true }));
  c.push(center("Disclosed UNDER SECTION 10(b)", { bold: true }));
  c.push(blank());
  c.push(plain("Prior Agreements:", { bold: true }));
  const priorAgmts = (data.priorAgreements || "").trim();
  if (priorAgmts) {
    priorAgmts.split(/\r?\n/).forEach(line => c.push(plain(line)));
  } else {
    c.push(plain("None."));
  }
  c.push(blank());
  c.push(plain("Competitors:", { bold: true }));
  const comps = (data.competitors || "").trim();
  if (comps) {
    comps.split(/\r?\n/).forEach(line => c.push(plain(line)));
  } else {
    c.push(plain("None."));
  }
  c.push(blank());
  c.push(plain(`Except as indicated above on this Exhibit D, Consultant has no disclosures to be made pursuant to Section 10(b) of this Agreement.`));
  c.push(blank());
  c.push(new Paragraph({ children: [new TextRun({ text: "CONSULTANT: ", bold: true }), new TextRun({ text: name })] }));
  c.push(blank());
  c.push(plain("By: __________________________"));
  c.push(plain(`Name: ${name}`));
  c.push(plain(`Address: ${addr}`));
  c.push(plain(`Email: ${email}`));
  c.push(plain(`Phone: ${phone}`));

  return new Document({
    styles: baseStyles,
    numbering: numberingConfig,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: c,
    }],
  });
}

// ============================================================
// EMPLOYEE DOC — placeholder baseline (swap with your real employee template)
// ============================================================
function buildEmployeeDoc(data) {
  const eff = formatDate(data.effectiveDate) || "[DATE]";
  const name = data.name || "[EMPLOYEE NAME]";
  const c = [];

  c.push(center(COMPANY.name, { bold: true, size: 28 }));
  c.push(blank());
  c.push(center("CONFIDENTIAL INFORMATION AND", { bold: true, size: 26 }));
  c.push(center("INVENTION ASSIGNMENT AGREEMENT", { bold: true, size: 26 }));
  c.push(blank());

  c.push(new Paragraph({ children: [
    new TextRun({ text: "Employee Name: ", bold: true }),
    new TextRun({ text: `${name} (\u201CEmployee\u201D)` }),
  ]}));
  c.push(new Paragraph({ children: [
    new TextRun({ text: "Effective Date: ", bold: true }),
    new TextRun({ text: eff }),
  ]}));
  c.push(blank());

  c.push(plain(`As a condition of becoming employed (or Employee\u2019s employment relationship being continued) by ${COMPANY.name}, a ${COMPANY.state} corporation (the \u201CCompany\u201D), and in consideration of Employee\u2019s employment with the Company and the compensation now and hereafter paid by the Company, Employee agrees to the following:`));

  c.push(numberedSection("At-Will Employment.", `Employee understands and acknowledges that Employee\u2019s employment with the Company is and shall continue to be at-will, meaning that either Employee or the Company may terminate the relationship at any time, with or without cause, and with or without notice.`));
  c.push(numberedSection("Confidential Information.", `Employee agrees to hold in strictest confidence, and not to use except for the benefit of the Company, any Confidential Information of the Company. This is a placeholder \u2014 replace the Employee tab body with your full employee CIIAA text.`));
  c.push(numberedSection("Assignment of Inventions.", `Employee hereby assigns to the Company all Company Inventions made or conceived during the term of employment.`));
  c.push(numberedSection("Return of Property.", `Upon termination, Employee will deliver all Company property and Confidential Information.`));
  c.push(numberedSection("Governing Law.", `Governed by the laws of the State of ${COMPANY.state}.`));

  c.push(blank());
  c.push(plain("THIS IS A PLACEHOLDER EMPLOYEE AGREEMENT. Replace buildEmployeeDoc() with your existing employee CIIAA text.", { italics: true, color: "B45309" }));
  c.push(blank());
  c.push(plain("THE COMPANY:", { bold: true }));
  c.push(plain(COMPANY.name, { bold: true }));
  c.push(plain("By: ___________________________"));
  c.push(plain(`        ${COMPANY.cfo}`));
  c.push(blank());
  c.push(new Paragraph({ children: [
    new TextRun({ text: "EMPLOYEE: ", bold: true }),
    new TextRun({ text: name }),
  ]}));
  c.push(plain("By: ____________________________"));

  return new Document({
    styles: baseStyles,
    numbering: numberingConfig,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: c,
    }],
  });
}

// ============================================================
// React UI
// ============================================================
export default function CIIAAGenerator() {
  const [tab, setTab] = useState("consultant");
  const [empData, setEmpData] = useState({ name: "", effectiveDate: "" });
  const [conData, setConData] = useState({
    name: "", effectiveDate: "", address: "", email: "", phone: "",
    priorInventions: "", priorAgreements: "", competitors: "",
  });
  const [busy, setBusy] = useState(false);

  const generate = async (kind) => {
    setBusy(true);
    try {
      const doc = kind === "employee" ? buildEmployeeDoc(empData) : buildConsultantDoc(conData);
      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const who = (kind === "employee" ? empData.name : conData.name) || (kind === "employee" ? "Employee" : "Consultant");
      const safeName = who.replace(/[^a-z0-9]+/gi, "_");
      a.href = url;
      a.download = `CIIAA_${kind === "employee" ? "Employee" : "Consultant"}_${safeName}.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (e) {
      console.error(e);
      alert("Generation failed: " + e.message);
    } finally {
      setBusy(false);
    }
  };

  const isConsultant = tab === "consultant";
  const accent = isConsultant ? "amber" : "blue";

  return (
    <div className="min-h-screen bg-stone-50 text-stone-900">
      <div className="max-w-4xl mx-auto p-6 md:p-10">
        {/* Header */}
        <header className="mb-8">
          <div className="flex items-center gap-3 mb-2">
            <div className="w-10 h-10 rounded-lg bg-stone-900 text-stone-50 flex items-center justify-center">
              <Building2 className="w-5 h-5" />
            </div>
            <div>
              <div className="text-xs uppercase tracking-widest text-stone-500">{COMPANY.name}</div>
              <h1 className="text-2xl font-semibold tracking-tight">CIIAA Generator</h1>
            </div>
          </div>
          <p className="text-sm text-stone-600 max-w-2xl">
            Generate a Confidential Information and Invention Assignment Agreement for an employee or a consultant. Each type uses different terms — choose carefully.
          </p>
        </header>

        {/* Tabs */}
        <div className="flex gap-2 mb-6 border-b border-stone-200">
          <button
            onClick={() => setTab("employee")}
            className={`px-4 py-3 text-sm font-medium border-b-2 -mb-px transition-colors flex items-center gap-2 ${
              tab === "employee" ? "border-blue-600 text-blue-700" : "border-transparent text-stone-500 hover:text-stone-800"
            }`}
          >
            <User className="w-4 h-4" /> Employee
          </button>
          <button
            onClick={() => setTab("consultant")}
            className={`px-4 py-3 text-sm font-medium border-b-2 -mb-px transition-colors flex items-center gap-2 ${
              tab === "consultant" ? "border-amber-600 text-amber-700" : "border-transparent text-stone-500 hover:text-stone-800"
            }`}
          >
            <Briefcase className="w-4 h-4" /> Consultant
          </button>
        </div>

        {/* Differentiator banner */}
        <div className={`rounded-lg p-4 mb-6 flex gap-3 border ${
          isConsultant ? "bg-amber-50 border-amber-200 text-amber-900" : "bg-blue-50 border-blue-200 text-blue-900"
        }`}>
          <AlertCircle className="w-5 h-5 flex-shrink-0 mt-0.5" />
          <div className="text-sm">
            {isConsultant ? (
              <>
                <strong>Consultant agreement.</strong> For non-employees (independent contractors). References a separate Consulting Agreement and covers Consultant\u2019s personnel. Includes Exhibit D (prior agreements + competitors) and a Notice to Third Parties clause. Not for W-2 employees.
              </>
            ) : (
              <>
                <strong>Employee agreement.</strong> For W-2 employees only. Includes at-will employment terms. For independent contractors, switch to the Consultant tab.
              </>
            )}
          </div>
        </div>

        {/* Form */}
        {tab === "employee" ? (
          <div className="bg-white rounded-xl border border-stone-200 p-6 space-y-4">
            <Field label="Employee Name" value={empData.name} onChange={(v) => setEmpData({ ...empData, name: v })} placeholder="Jane Smith" />
            <Field label="Effective Date" type="date" value={empData.effectiveDate} onChange={(v) => setEmpData({ ...empData, effectiveDate: v })} />
            <button
              disabled={busy}
              onClick={() => generate("employee")}
              className="mt-2 inline-flex items-center gap-2 px-4 py-2.5 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 disabled:opacity-50"
            >
              <Download className="w-4 h-4" />
              {busy ? "Generating\u2026" : "Generate Employee CIIAA"}
            </button>
          </div>
        ) : (
          <div className="bg-white rounded-xl border border-stone-200 p-6 space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <Field label="Consultant Name" value={conData.name} onChange={(v) => setConData({ ...conData, name: v })} placeholder="John Doe" />
              <Field label="Effective Date" type="date" value={conData.effectiveDate} onChange={(v) => setConData({ ...conData, effectiveDate: v })} />
            </div>
            <Field label="Address" value={conData.address} onChange={(v) => setConData({ ...conData, address: v })} placeholder="123 Main St, City, State ZIP" />
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <Field label="Email" type="email" value={conData.email} onChange={(v) => setConData({ ...conData, email: v })} placeholder="john@example.com" />
              <Field label="Phone" value={conData.phone} onChange={(v) => setConData({ ...conData, phone: v })} placeholder="(555) 555-5555" />
            </div>

            <div className="pt-4 border-t border-stone-100 space-y-4">
              <div className="text-xs uppercase tracking-wider text-stone-500 font-medium">Exhibits (one item per line, leave blank for \u201CNone\u201D)</div>
              <TextArea label="Exhibit A \u2014 Prior Inventions" value={conData.priorInventions} onChange={(v) => setConData({ ...conData, priorInventions: v })} placeholder="Describe each prior invention on its own line" />
              <TextArea label="Exhibit D \u2014 Prior Agreements" value={conData.priorAgreements} onChange={(v) => setConData({ ...conData, priorAgreements: v })} placeholder="e.g. NDA with Acme Corp (2023)" />
              <TextArea label="Exhibit D \u2014 Competitors" value={conData.competitors} onChange={(v) => setConData({ ...conData, competitors: v })} placeholder="Companies you currently work with that may compete with Skyfire" />
            </div>

            <button
              disabled={busy}
              onClick={() => generate("consultant")}
              className="mt-2 inline-flex items-center gap-2 px-4 py-2.5 bg-amber-600 text-white rounded-lg font-medium hover:bg-amber-700 disabled:opacity-50"
            >
              <Download className="w-4 h-4" />
              {busy ? "Generating\u2026" : "Generate Consultant CIIAA"}
            </button>
          </div>
        )}

        <footer className="mt-8 text-xs text-stone-500 flex items-center gap-2">
          <FileText className="w-3.5 h-3.5" />
          Output is a .docx file. Review before signing.
        </footer>
      </div>
    </div>
  );
}

function Field({ label, value, onChange, type = "text", placeholder }) {
  return (
    <label className="block">
      <div className="text-xs font-medium text-stone-600 mb-1.5">{label}</div>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="w-full px-3 py-2 border border-stone-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-stone-400 focus:border-stone-400"
      />
    </label>
  );
}

function TextArea({ label, value, onChange, placeholder }) {
  return (
    <label className="block">
      <div className="text-xs font-medium text-stone-600 mb-1.5">{label}</div>
      <textarea
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        rows={3}
        className="w-full px-3 py-2 border border-stone-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-stone-400 focus:border-stone-400 resize-y"
      />
    </label>
  );
}
