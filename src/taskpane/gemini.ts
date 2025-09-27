import { GoogleGenerativeAI } from "@google/generative-ai";

// eslint-disable-next-line no-undef
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY!);
const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash-lite" });

export async function getFormattedReferences(references: string): Promise<string> {
  const prompt = `
    Rewrite the following list of references like this: (author name, year). Put blank space between references and write them one below the other, also don't number them or write them in bullet points. Never write something like "Here is the list of references formatted in the style you requested:" or anything close to that. If the list of references include the following words, never add them in the output text: Books, Journal Articles, Websites, Additional References, Remember to write the references between parenthesis like this: (author name, year)  -> (Buzan, 2010). Write citations under the format (author name, year), not other descriptions. 
    Don't forget to put blank space after each reference. 
    Example:
    \`\`\`
    (Buzan, 2010)
    
    (Buzan, 2010)
    \`\`\`
    Here's the list:
    ${references}
    `;

  const result = await model.generateContent(prompt);
  const response = result.response;
  return response.text();
}
