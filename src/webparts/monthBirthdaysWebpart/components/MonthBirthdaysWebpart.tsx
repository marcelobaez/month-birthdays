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
import { endOfMonth, format, setYear } from "date-fns";
// import { setYear } from "date-fns";
import {
  ActionButton,
  Callout,
  getId,
  Icon,
  IStackTokens,
  Link,
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
const dateFormat: string = "dd-MM-yyyy";
const today: string = format(pivotDay, dateFormat);

interface ICell {
  Key: string;
  Value: string;
}

interface IPerson {
  Cells: ICell[];
}

interface ISearchResult {
  PrimaryQueryResult: {
    RelevantResults: {
      Table: {
        Rows: IPerson[];
      };
    };
  };
}

export interface IState {
  monthBirthdays?: IPerson[];
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
    const parsedCurrent: string = format(new Date(currentDate), dateFormat);
    return parsedCurrent === today;
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
                  ¿Qué hacer si mi cumpleaños no aparece aquí?
                </p>
              </div>
              <div className={mainStyles.inner}>
                <div>
                  <p className={mainStyles.subtext} id={this._descriptionId}>
                    Es probable que no haya completado la fecha de su nacimiento
                    en su perfil de Office.
                  </p>
                </div>
                <div className={mainStyles.actions}>
                  <Link
                    className={mainStyles.link}
                    href="https://ebyorgar-my.sharepoint.com/_layouts/15/me.aspx?v=profile"
                    target="_blank"
                  >
                    Ir a mi perfil
                  </Link>
                </div>
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
          monthBirthdays.map((person: IPerson) => {
            const birthday: string =
              person.Cells.find((cell: ICell) => cell.Key === "Birthday5")
                ?.Value || "";
            const fullName: string =
              person.Cells.find((cell: ICell) => cell.Key === "PreferredName")
                ?.Value || "";
            const officeNumber: string =
              person.Cells.find((cell: ICell) => cell.Key === "OfficeNumber")
                ?.Value || "";
            const pictureUrl: string =
              person.Cells.find((cell: ICell) => cell.Key === "PictureUrl")
                ?.Value || "";
            const showToday: boolean = this.compareToToday(birthday);

            const examplePersona: IPersonaSharedProps = {
              imageUrl: pictureUrl,
              text: fullName,
              secondaryText: officeNumber,
              tertiaryText: format(new Date(birthday), "dd/MM"),
              optionalText:
                person.Cells.find((cell: ICell) => cell.Key === "AADObjectID")
                  ?.Value || "",
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
                key={person.Cells[1].Value}
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
    const { context, maxDisplayNumber } = this.props;
    const lastDayOfMonth: string = format(endOfMonth(pivotDay), "yyyy-MM-dd");

    try {
      this.setState({ IsLoading: true });
      const queryMonth: SPHttpClientResponse = await context.spHttpClient.get(
        context.pageContext.web.absoluteUrl +
          "/_api/search/query?querytext='Birthday5<=\"" +
          lastDayOfMonth +
          '" AND Birthday5>="' +
          format(pivotDay, "yyyy-MM-dd") +
          "\"'" +
          "&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'" +
          "&selectproperties='PreferredName,FirstName,LastName,Birthday5," +
          "PictureUrl,OfficeNumber,WorkPhone,AADObjectID'" +
          "&rowlimit=" +
          maxDisplayNumber.toString() +
          "&sortlist='Birthday5:descending'",
        clientConfigODataV3
      );

      const resultsMonth: ISearchResult = await queryMonth.json();

      if (
        resultsMonth.PrimaryQueryResult.RelevantResults.Table.Rows.length === 0
      ) {
        this.setState({ IsDataFound: false });
      }

      this.setState({
        IsLoading: false,
        IsDataFound: true,
        monthBirthdays:
          resultsMonth.PrimaryQueryResult.RelevantResults.Table.Rows,
      });
    } catch (err) {
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
        <Link
          target="_blank"
          href={`https://lam.delve.office.com/?u=${props.optionalText}&v=work`}
        >
          {"Ir al perfil"}
        </Link>
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
