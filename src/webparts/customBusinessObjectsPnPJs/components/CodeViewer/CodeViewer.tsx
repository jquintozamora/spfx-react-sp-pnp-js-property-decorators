import * as React from "react";

import SyntaxHighlighter, { registerLanguage } from "react-syntax-highlighter/dist/light"
import ts from 'react-syntax-highlighter/dist/languages/typescript';
import docco from 'react-syntax-highlighter/dist/styles/docco';
registerLanguage('typescript', ts);

export type Props = {
  className?: string,
  style?: React.CSSProperties,
  code?: string
};

const CodeViewer: React.StatelessComponent<Props> = (props) => {
  const { code } = this.props;
  return (
    <div>
      <SyntaxHighlighter language='typescript' style={docco}>{code}</SyntaxHighlighter>
      <div>
        <button onClick={this._firstCode}>Play</button>
      </div>
    </div>
  );
};

export default CodeViewer;
