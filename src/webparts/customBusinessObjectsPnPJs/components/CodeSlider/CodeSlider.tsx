import * as React from "react";
import { Tabs, Tab } from "pui-react-tabs";
import CodeViewer from "../CodeViewer/CodeViewer";

import styles from "./CodeSlider.module.scss";

import { ICodeSliderProps } from "./ICodeSliderProps";
import { ICodeSliderState } from "./ICodeSliderState";

import { samples } from "../../data/Samples";

export default class CodeSlider extends React.Component<ICodeSliderProps, ICodeSliderState> {

  constructor(props: ICodeSliderProps) {
    super(props);
    // set initial state
    this.state = {
      codeSamples: samples,
      errors: [],
    };
  }

  public render(): React.ReactElement<ICodeSliderProps> {
    return (
      <div className={"ms-bgColor-themeDark " + styles.container}>
        <span className="ms-font-xl ms-fontColor-white">{this.props.title}</span>
        <Tabs defaultActiveKey={1} >
          {this._getTabs()}
        </Tabs>
        <div>
          {this._gerErrors()}
        </div>
      </div >
    );
  }

  private _getTabs() {
    return this.state.codeSamples.map((sample, index) => {
      return (
        <Tab eventKey={index + 1} title={sample.tabTitle} tabClassName={styles.tab} className={styles.tabContent} >
          <CodeViewer
            codeText={sample.codeText()}
            codeFunc={sample.code}
            title={sample.title}
          />
        </Tab>
      );
    });
  }

  private _gerErrors() {
    return this.state.errors.length > 0
      ?
      <div style={{ color: "orangered" }} >
        <div>Errors:</div>
        {
          this.state.errors.map((item) => {
            return (<div>{JSON.stringify(item)}</div>);
          })
        }
      </div>
      : null;
  }

}
