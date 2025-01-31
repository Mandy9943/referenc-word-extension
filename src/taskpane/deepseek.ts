import OpenAI from "openai";

const deepseek = new OpenAI({
  // eslint-disable-next-line no-undef
  baseURL: process!.env.API_URL!,
  // eslint-disable-next-line no-undef
  apiKey: process!.env.API_KEY!,
  dangerouslyAllowBrowser: true,
});

export async function getFormattedReferences(references: string): Promise<string> {
  const prompt = `
    Rewrite the following list of references like this: (author name, year). Put blank space between references and write them one below the other, also don’t number them or write them in bullet points. Never write something like “Here is the list of references formatted in the style you requested:” or anything close to that. If the list of references include the following words, never add them in the output text: Books, Journal Articles, Websites, Additional References, Remember to write the references between parenthesis like this: (author name, year)  -> (Buzan, 2010). Write citations under the format (author name, year), not other descriptions. 
    Don’t forget to put blank space after each reference. Here’s the list:
    ${references}
    `;

  const completion = await deepseek.chat.completions.create({
    messages: [
      {
        role: "system",
        content: "You are a helpful assistant that formats references in a specific way.",
      },
      {
        role: "user",
        content: prompt,
      },
    ],
    model: "deepseek-chat",
  });
  return completion.choices[0].message.content;
}
