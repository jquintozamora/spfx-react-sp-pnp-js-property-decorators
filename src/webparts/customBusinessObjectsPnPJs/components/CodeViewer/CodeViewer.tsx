import * as React from "react";

import styles from "./CodeViewer.module.scss";

import SyntaxHighlighter, { registerLanguage } from "react-syntax-highlighter/dist/light"
import ts from 'react-syntax-highlighter/dist/languages/typescript';
import docco from 'react-syntax-highlighter/dist/styles/docco';
registerLanguage('typescript', ts);

export type Props = {
  className?: string,
  style?: React.CSSProperties,
  title: string,
  codeText: string,
  codeFunc: any
};

const CodeViewer: React.StatelessComponent<Props> = (props) => {
  const { codeText, codeFunc, title } = props;
  return (
    <div>
      <h2 className={styles.title}>{title}</h2>
      <SyntaxHighlighter language='typescript' style={docco}>{codeText}</SyntaxHighlighter>
      <div className={styles.run}>
        <button  onClick={codeFunc}>Run Code</button>
      </div>
    </div>
  );
};

export default CodeViewer;
