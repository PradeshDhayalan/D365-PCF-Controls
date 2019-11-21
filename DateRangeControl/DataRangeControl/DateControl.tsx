import * as React from "react";
import * as ReactDOM from "react-dom";
import { DateRange,CalendarTheme } from 'react-date-range'
import { CANCELLED } from "dns";
import { Moment }from 'moment';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Callout, DirectionalHint } from 'office-ui-fabric-react/lib/Callout';
import { TextField } from 'office-ui-fabric-react';
import { FocusTrapZone } from 'office-ui-fabric-react/lib/FocusTrapZone';

export interface IPCFDateRangeControlState 
    extends React.ComponentState,IPCFDateRangeControlProps { }

export interface IPCFDateRangeControlProps {
    showCalendar: boolean;
    selectedStartDate: Moment | undefined;
    selectedEndDate: Moment | undefined;
    textFieldValue: string,
    startDisplayName: string,
    endDisplayName: string,
    inputDateChanged?: (dtStart: Moment | undefined, dtEndDate: Moment | undefined) => void;
}

export class PCFDateRangeControl extends React.Component<IPCFDateRangeControlProps,IPCFDateRangeControlState>
{
    parseDtText(selectedStartDate: Moment | undefined, selectedEndDate: Moment | undefined, textFieldValue: string): string {
        return (textFieldValue == "" && selectedEndDate != undefined && selectedStartDate != undefined) ? this.parseDateTime(this.convertToDate(selectedStartDate)) + " - " + this.parseDateTime(this.convertToDate(selectedEndDate)) : textFieldValue;
    }
    private _calendarButtonElement!: HTMLElement;

    constructor(props:IPCFDateRangeControlProps) {
        super(props);

        this.state = {
            showCalendar : props.showCalendar,
            selectedStartDate: props.selectedStartDate,
            selectedEndDate: props.selectedEndDate,
            textFieldValue: this.parseDtText(props.selectedStartDate,props.selectedEndDate,props.textFieldValue),
            startDisplayName: props.startDisplayName,
            endDisplayName: props.endDisplayName,
        };
    }

    private _onClick(event: any): void {
        this.setState((prevState: IPCFDateRangeControlState) => {
            prevState.showCalendar = !prevState.showCalendar;
            return prevState;
        });
    }

    private _onDismiss(event: any): void {
        this.setState((prevState: IPCFDateRangeControlState) => {
            prevState.showCalendar = false;
            return prevState;
        });
    }

    private convertToDate(date: any): Date {
        return date.toDate() as Date;
    }

    private handleRangeChange(which: any, payload : any) {

        this.setState((prevState: IPCFDateRangeControlState) => {
            prevState.selectedStartDate = payload.startDate as Moment;
            prevState.selectedEndDate = payload.endDate as Moment;
            prevState.textFieldValue = this.parseDateTime(payload.startDate.toDate()) + " - " + this.parseDateTime(payload.endDate.toDate());
            return prevState;
        });

        if(this.props.inputDateChanged) {
            this.props.inputDateChanged(payload.startDate as Moment,
                payload.endDate as Moment);
        }
    }

    parseDateTime(dt: Date): string {
        
        var day = (dt.getDate());
        var month_index = dt.getMonth() + 1;
        var year = dt.getFullYear();
        
        return day + "/" + month_index + "/" + year;
      }

    public render() : JSX.Element {

        initializeIcons();

        return(
            <div>

            <div ref={calendarBtn => ( this._calendarButtonElement = calendarBtn!)}>
                <TextField readOnly value={ this.state.textFieldValue } onClick={ this._onClick.bind(this) } iconProps={{ iconName: 'Calendar'}} />
            </div>
            <div>
            { this.state.showCalendar && (<Callout
                    isBeakVisible={false}
                    className="ms-DatePicker-callout"
                    gapSpace={0}
                    doNotLayer={false}
                    target={this._calendarButtonElement}
                    directionalHint={DirectionalHint.bottomLeftEdge}
                    onDismiss = {this._onDismiss.bind(this)}
                    setInitialFocus={false}>
                    <FocusTrapZone isClickableOutsideFocusTrap={true}>
                    <DateRange startDate={ this.state.selectedStartDate } endDate={ this.state.selectedEndDate }  onChange={this.handleRangeChange.bind(this,'dateRangePicker') }/>
                    </FocusTrapZone>
            </Callout> )}   
            </div>
            </div>
        )
    }
}

export default PCFDateRangeControl ;