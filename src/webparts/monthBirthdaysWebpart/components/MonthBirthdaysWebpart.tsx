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
import { FixedSizeList } from "react-window";
import { FoodCakeFilled, EditRegular } from "@fluentui/react-icons";
import {
  webLightTheme,
  Button,
  FluentProvider,
  List,
  ListItem,
  MessageBar,
  MessageBarTitle,
  Persona,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Spinner,
} from "@fluentui/react-components";
import {
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
interface IOdataResponse {
  "odata.metadata": string;
  "odata.nextLink": string;
  value: {
    Title: string;
    field_1: string;
    "odata.etag": string;
    "odata.editLink": string;
    "odata.id": string;
    "odata.type": string;
  }[];
}

interface IMonthBirthday {
  Title: string;
  field_1: string;
}

export interface IState {
  monthBirthdays?: IMonthBirthday[];
  IsLoading?: boolean;
  IsDataFound?: boolean;
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
    };
  }

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
    const [day, month] = currentDate.split("/");
    const normalizedDay = day.length === 1 ? `0${day}` : day;
    const normalizedMonth = month.length === 1 ? `0${month}` : month;
    const normalizedDate = `${normalizedDay}/${normalizedMonth}`;
    return normalizedDate === format(pivotDay, "dd/MM");
  }

  public BirthdayList = React.forwardRef<HTMLUListElement>(
    (props: React.ComponentProps<typeof List>, ref) => (
      <List aria-label="Countries" tabIndex={0} {...props} ref={ref} />
    )
  );

  public render(): React.ReactElement<ITestWebpartNode22Props> {
    const { monthBirthdays, IsLoading, IsDataFound } = this.state;

    return (
      <FluentProvider theme={webLightTheme}>
        <div style={{ display: "flex", flexDirection: "column", gap: "20px" }}>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
            }}
          >
            <div>
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
                  <span style={webpartHeaderViewStyles}>
                    {this.props.title}
                  </span>
                )}
              </div>
            </div>
            <Popover>
              <PopoverTrigger disableButtonEnhancement>
                <Button appearance="subtle" icon={<EditRegular />}>
                  ¿Querés modificar alguna información?
                </Button>
              </PopoverTrigger>
              <PopoverSurface tabIndex={-1}>
                <>
                  <h3>Contactanos</h3>
                  <div>
                    Si preferís no mostrar tu cumpleaños o detectás un error,
                    podés escribirnos a rrhhbue@eby.org.ar
                  </div>
                </>
              </PopoverSurface>
            </Popover>
          </div>
          {IsLoading && <Spinner size="large" label="Buscando cumpleaños..." />}
          {!IsLoading &&
          IsDataFound &&
          monthBirthdays &&
          monthBirthdays.length > 0 ? (
            <FixedSizeList
              height={300}
              itemCount={monthBirthdays.length}
              itemSize={50}
              width="100%"
              itemData={monthBirthdays}
              outerElementType={this.BirthdayList}
            >
              {({ index, style, data }) => {
                const person = data[index];
                const birthday: string = person.field_1;
                const fullName: string = person.Title;
                const showToday: boolean = this.compareToToday(birthday);

                return (
                  <ListItem
                    style={style}
                    aria-setsize={monthBirthdays.length}
                    aria-posinset={index + 1}
                  >
                    <Persona
                      avatar={{ color: "colorful" }}
                      name={fullName}
                      size={showToday ? "extra-large" : "medium"}
                      presence={undefined}
                      secondaryText={
                        <React.Fragment>
                          <FoodCakeFilled style={{ marginRight: "5px" }} />
                          {showToday ? "¡Hoy cumple años!" : birthday}
                        </React.Fragment>
                      }
                    />
                  </ListItem>
                );
              }}
            </FixedSizeList>
          ) : (
            <MessageBar>
              <MessageBarTitle>No hay cumpleaños este mes :(.</MessageBarTitle>
            </MessageBar>
          )}
        </div>
      </FluentProvider>
    );
  }

  private async _getListData(): Promise<void> {
    const { context } = this.props;

    try {
      this.setState({ IsLoading: true });

      // Use an OR expression in the OData filter to check for both '/06' and '/6'
      const endpoint = `https://ebyorgar.sharepoint.com/sites/intranet_EBY/_api/web/lists/getbytitle('birthdays_eby')/items?$select=Title,field_1&$top=2000`;
      const response: SPHttpClientResponse = await context.spHttpClient.get(
        endpoint,
        clientConfigODataV3
      );
      const data: IOdataResponse = await response.json();

      const today = new Date();
      const todayDay = today.getDate();
      const todayMonth = today.getMonth() + 1;

      const filtered = data.value
        .filter((item: IMonthBirthday) => {
          // Only keep items for the current month
          const [dayStr, monthStr] = item.field_1
            .split("/")
            .map((s) => s.trim());
          const day = parseInt(dayStr, 10);
          const month = parseInt(monthStr, 10);

          // Only birthdays in the current month
          if (month !== todayMonth) return false;

          // Only birthdays today or later
          return day >= todayDay;
        })
        .sort((a: IMonthBirthday, b: IMonthBirthday) => {
          const [dayA] = a.field_1.split("/").map(Number);
          const [dayB] = b.field_1.split("/").map(Number);
          return dayA - dayB;
        });

      const resultsMonth = filtered.map((item: IMonthBirthday) => {
        return {
          Title: item.Title,
          field_1: item.field_1,
        };
      });

      if (resultsMonth.length === 0) {
        this.setState({ IsDataFound: false });
      }

      this.setState({
        IsLoading: false,
        IsDataFound: true,
        monthBirthdays: resultsMonth,
      });
    } catch (err) {
      console.log(err);
      console.error("Ocurrió un error:", err.message);
    }
  }

  private setTitle(event: React.FormEvent<HTMLTextAreaElement>): void {
    this.props.setTitle(event.currentTarget.value);
  }
}
