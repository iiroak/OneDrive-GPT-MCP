const TOOL_OUTPUT_SCHEMA = {
  type: 'object',
  properties: {
    message: {
      type: 'string',
      description: 'Human-readable result of the tool call.'
    },
    data: {
      description: 'Optional structured details returned by the tool.'
    }
  },
  required: ['message'],
  additionalProperties: false
};

function toStructuredContent(result) {
  const message = Array.isArray(result?.content)
    ? result.content
      .filter(item => item?.type === 'text' && typeof item.text === 'string')
      .map(item => item.text)
      .join('\n')
    : '';
  const structured = { message: message || 'Tool completed successfully.' };

  if (result && Object.prototype.hasOwnProperty.call(result, 'structuredContent')) {
    structured.data = result.structuredContent;
  }

  return structured;
}

module.exports = { TOOL_OUTPUT_SCHEMA, toStructuredContent };
