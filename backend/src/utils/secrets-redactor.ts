const SECRET_PATTERNS = [
  /om-[a-zA-Z0-9]{32,}/g,
  /sk-[a-zA-Z0-9]{32,}/g,
  /sk-ant-[a-zA-Z0-9]{32,}/g,
  /gsk_[a-zA-Z0-9]{32,}/g,
  /gig_[a-zA-Z0-9]{32,}/g,
];

export function redactSecrets(text: string): string {
  let result = text;
  for (const pattern of SECRET_PATTERNS) {
    result = result.replace(pattern, "***");
  }
  return result;
}
