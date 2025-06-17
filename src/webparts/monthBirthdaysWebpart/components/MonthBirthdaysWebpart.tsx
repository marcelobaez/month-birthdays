import * as React from "react";
import type { ITestWebpartNode22Props } from "./IMonthBirthdaysWebpartProps";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ODataVersion,
  ISPHttpClientConfiguration,
  SPHttpClientConfiguration,
} from "@microsoft/sp-http";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { format, setYear } from "date-fns";
import {
  ActionButton,
  Callout,
  getId,
  Icon,
  IStackTokens,
  PersonaSize,
  Spinner,
  SpinnerSize,
  Stack,
} from "@fluentui/react";
import { IPersonaProps, IPersonaSharedProps } from "@fluentui/react";
import { PersonaPresence } from "@fluentui/react";
import { Persona } from "@fluentui/react";
import { MessageBar } from "@fluentui/react";
import {
  mainStyles,
  webpartHeaderEditStyles,
  webpartHeaderStyles,
  webpartHeaderViewStyles,
} from "../utils";

// create a new ISPHttpClientConfiguration object with defaultODataVersion = ODataVersion.v3
const spSearchConfig: ISPHttpClientConfiguration = {
  defaultODataVersion: ODataVersion.v3,
};

// sobre-escribir ODataVersion.v4 flag with the ODataVersion.v3
const clientConfigODataV3: SPHttpClientConfiguration =
  SPHttpClient.configurations.v1.overrideWith(spSearchConfig);

const pivotDay: Date = setYear(new Date(), 2000);

  interface IOdataResponse { "odata.metadata": string; "odata.nextLink": string; value: { Title: string; field_1: string; "odata.etag": string; "odata.editLink": string; "odata.id": string; "odata.type": string }[] }

interface IMonthBirthday {
  Title: string;
  field_1: string;
}

export interface IState {
  monthBirthdays?: IMonthBirthday[];
  IsLoading?: boolean;
  IsDataFound?: boolean;
  IsCalloutVisible?: boolean;
}

export default class TestWebpartNode22 extends React.Component<
  ITestWebpartNode22Props,
  IState
> {
  constructor(props: ITestWebpartNode22Props) {
    super(props);
    this.state = {
      monthBirthdays: [],
      IsLoading: false,
      IsDataFound: false,
      IsCalloutVisible: false,
    };
  }

  private _menuButtonElement = React.createRef<HTMLDivElement>();
  // Use getId() to ensure that the callout label and description IDs are unique on the page.
  // (It's also okay to use plain strings without getId() and manually ensure their uniqueness.)
  private _labelId: string = getId("callout-label");
  private _descriptionId: string = getId("callout-description");

  public async componentDidMount(): Promise<void> {
    if (Environment.type === EnvironmentType.SharePoint) {
      await this._getListData();
    }
  }

  public async componentDidUpdate(
    prevProps: ITestWebpartNode22Props,
    prevState: IState
  ): Promise<void> {
    if (prevProps !== this.props) {
      await this._getListData();
    }
  }

  public compareToToday(currentDate: string): boolean {
    const [day, month] = currentDate.split('/');
    const normalizedDay = day.length === 1 ? `0${day}` : day;
    const normalizedMonth = month.length === 1 ? `0${month}` : month;
    const normalizedDate = `${normalizedDay}/${normalizedMonth}`;
    return normalizedDate === format(pivotDay, "dd/MM");
  }

  public render(): React.ReactElement<ITestWebpartNode22Props> {
    const { monthBirthdays, IsLoading, IsDataFound, IsCalloutVisible } =
      this.state;

    const containerStackTokens: IStackTokens = { childrenGap: 20 };

    return (
      <Stack tokens={containerStackTokens}>
        <Stack horizontal horizontalAlign="space-between">
          <Stack.Item align="start">
            <div style={webpartHeaderStyles}>
              {this.props.isEditMode && (
                <textarea
                  onChange={this.setTitle.bind(this)}
                  style={webpartHeaderEditStyles}
                  placeholder="Agregar un título"
                  aria-label="Agregar un título"
                  defaultValue={this.props.title}
                />
              )}
              {!this.props.isEditMode && (
                <span style={webpartHeaderViewStyles}>{this.props.title}</span>
              )}
            </div>
          </Stack.Item>
          <Stack.Item align="end">
            <span ref={this._menuButtonElement}>
              <ActionButton
                data-automation-id="test"
                iconProps={{ iconName: "EmojiNeutral" }}
                allowDisabledFocus={true}
                onClick={this._onShowMenuClicked}
              >
                ¿No apareces?
              </ActionButton>
            </span>
            <Callout
              className={mainStyles.callout}
              ariaLabelledBy={this._labelId}
              ariaDescribedBy={this._descriptionId}
              role="alertdialog"
              gapSpace={0}
              target={this._menuButtonElement.current}
              onDismiss={this._onCalloutDismiss}
              setInitialFocus={true}
              hidden={!IsCalloutVisible}
            >
              <div className={mainStyles.header}>
                <p className={mainStyles.title} id={this._labelId}>
                Si preferís no mostrar tu cumpleaños o detectás un error,
                podés escribirnos a rrhhbue@eby.org.ar
                </p>
              </div>
            </Callout>
          </Stack.Item>
        </Stack>
        {IsLoading && (
          <Spinner size={SpinnerSize.large} label="Buscando cumpleaños..." />
        )}
        {!IsLoading &&
        IsDataFound &&
        monthBirthdays &&
        monthBirthdays.length > 0 ? (
          monthBirthdays.map((person: IMonthBirthday) => {
            const birthday: string =
              person.field_1;
            const fullName: string =
              person.Title;
            const showToday: boolean = this.compareToToday(birthday);

            const examplePersona: IPersonaSharedProps = {
              text: fullName,
              tertiaryText: birthday,
            };

            return (
              <Persona
                {...examplePersona}
                size={showToday ? PersonaSize.size72 : PersonaSize.size48}
                presence={PersonaPresence.none}
                onRenderSecondaryText={(props: IPersonaProps) =>
                  this._onRenderSecondaryText(
                    props.secondaryText,
                    props.tertiaryText,
                    showToday
                  )
                }
                onRenderTertiaryText={this._onRenderTertiaryText}
                hidePersonaDetails={false}
                key={person.field_1 + person.Title}
              />
            );
          })
        ) : (
          <MessageBar>No hay cumpleaños este mes :(.</MessageBar>
        )}
      </Stack>
    );
  }

  private async _getListData(): Promise<void> {
    const { context } = this.props;

    try {
      this.setState({ IsLoading: true });

      // Use an OR expression in the OData filter to check for both '/06' and '/6'
      const endpoint = `https://ebyorgar.sharepoint.com/sites/intranet_EBY/_api/web/lists/getbytitle('birthdays_eby')/items?$select=Title,field_1`;
      const response: SPHttpClientResponse = await context.spHttpClient.get(endpoint, clientConfigODataV3);
      const data: IOdataResponse = await response.json();

      const today = new Date();
      const todayDay = today.getDate();
      const todayMonth = today.getMonth() + 1;

      const filtered = data.value
        .filter((item: IMonthBirthday) => {
          // Only keep items for the current month
          const [dayStr, monthStr] = item.field_1.split('/');
          const day = parseInt(dayStr, 10);
          const month = parseInt(monthStr, 10);

          // Only birthdays in the current month
          if (month !== todayMonth) return false;

          // Only birthdays today or later
          return day >= todayDay;
        })
        .sort((a: IMonthBirthday, b: IMonthBirthday) => {
          const [dayA] = a.field_1.split('/').map(Number);
          const [dayB] = b.field_1.split('/').map(Number);
          return dayA - dayB;
        });
      
      const resultsMonth = filtered.map((item: IMonthBirthday) => {
        return {
          Title: item.Title,
          field_1: item.field_1,
        };
      });

      if (
        resultsMonth.length === 0
      ) {
        this.setState({ IsDataFound: false });
      }

      this.setState({
        IsLoading: false,
        IsDataFound: true,
        monthBirthdays:
          resultsMonth,
      });
    } catch (err) {
      console.log(err);
      console.error("Ocurrió un error:", err.message);
    }
  }

  private _onRenderSecondaryText = (
    secondaryText: string | undefined,
    tertiaryText: string | undefined,
    showToday: boolean
  ): JSX.Element => {
    return (
      <div>
        {secondaryText && (
          <React.Fragment>
            <Icon iconName={"Suitcase"} style={{ marginRight: "5px" }} />
            {secondaryText}
          </React.Fragment>
        )}
        {!showToday && (
          <React.Fragment>
            <Icon
              iconName={"BirthdayCake"}
              style={{
                marginRight: "5px",
                marginLeft: secondaryText ? "10px" : "0px",
                color: "#5C2E91",
              }}
            />
            {tertiaryText}
          </React.Fragment>
        )}
      </div>
    );
  };

  private _onRenderTertiaryText = (props: IPersonaProps): JSX.Element => {
    return (
      <div>
        <Icon
          iconName={"BirthdayCake"}
          style={{
            marginRight: "5px",
            color: "#5C2E91",
          }}
        />
        {"¡Hoy cumple años! "}
      </div>
    );
  };

  private setTitle(event: React.FormEvent<HTMLTextAreaElement>): void {
    this.props.setTitle(event.currentTarget.value);
  }

  private _onShowMenuClicked = (): void => {
    this.setState({
      IsCalloutVisible: !this.state.IsCalloutVisible,
    });
  };

  private _onCalloutDismiss = (): void => {
    this.setState({
      IsCalloutVisible: false,
    });
  };
}
