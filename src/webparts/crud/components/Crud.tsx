import * as React from 'react';
import { TextField, Label, Panel, PrimaryButton, DetailsList, IconButton, IIconProps, PanelType, Dialog, DialogFooter, DialogType } from "@fluentui/react";
import { Web } from "sp-pnp-js";
import { ICrudProps } from './ICrudProps';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import styles from './Crud.module.scss';

export interface IEmployee {
  ID: number;
  Title: string;
  EmployeeName: string | undefined;
  Ename?: string | null
}

// const style = mergeStyleSets({
//   btn:{
//     margin:10,
//     padding:5
//   }
// })

const dialogProps = {
  type: DialogType.normal,
  title: 'Successful!',
  closeButtonArialLabel: 'Close',
  subText: 'Hi! How are you?'
}

const editIcon: IIconProps = { iconName: 'Edit' }
const deleteIcon: IIconProps = { iconName: 'Delete' }
const addIcon: IIconProps = { iconName: 'Add' }

const webURL = "https://cubicdirect.sharepoint.com/sites/Pavan";
const CRUDtwo: React.FC<ICrudProps> = () => {
  const [employees, setEmployees] = React.useState<IEmployee[]>([]);
  const [selectedEmployee, setSelectedEmployee] = React.useState<IEmployee | null>(null);
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [selectedPeople, setSelectedPeople] = React.useState<IPersonaProps[]>([]);
  const [isEditMode, setIsEditMode] = React.useState(false);
  const [errorMessage, setErrorMessage] = React.useState<string>('');
  const [openDialog, setOpenDialog] = React.useState(true)

  const fetchData = async (): Promise<void> => {
    const web = new Web(webURL);
    const list = web.lists.getByTitle("Example");
    try {
      const items = await list.items.select("ID", "EmployeeNameId", "Title").expand().getAll();
      console.log("Retreived Data--", items)

      const userPromises = items.map(async (item: any) => {
        if (item.EmployeeNameId) {
          try {
            // const single = await web.siteUsers.getById(item.EmployeeNameId).get();
            // console.log(single);

            const userDetails = await web.siteUsers.getById(item.EmployeeNameId).get();
            console.log(item.Ename);
            item.Ename = userDetails.Title
            console.log("Retrieve UD", userDetails);

          } catch (userError) {
            console.error("Error fetching user details:", userError);
          }
        }
        return item;
      });

      const itemsWithUserNames = await Promise.all(userPromises);
      setEmployees(itemsWithUserNames);
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  };

  React.useEffect(() => {
    fetchData().then((data) => console.log(data)).catch((error) => console.log(error));
  }, []);


  const createData = async (): Promise<void> => {
    const web = new Web(webURL);
    const list = web.lists.getByTitle("Example");

    if (!selectedEmployee) {
      console.error("Selected employee is null");
      return;
    }

    const userValue: string | undefined = selectedPeople.length > 0 ? String(selectedPeople[0].key) : undefined;

    const data: any = {
      Title: selectedEmployee?.Title || '',
      EmployeeNameId: { results: userValue ? [userValue] : [] },
    };

    console.log("Selected Data--", data);

    try {
      await list.items.add(data);
      // alert("Created Successfully");
      setErrorMessage('');
      setSelectedEmployee(null);
      await fetchData();
      setIsPanelOpen(false);
      setIsEditMode(false);
    } catch (error) {
      console.error("Error creating data:", error);
    }
  };


  const updateData = async (): Promise<void> => {
    if (!selectedEmployee) {
      return;
    }
    const web = new Web(webURL);
    const list = web.lists.getByTitle("Example");

    let alluser: any = [];
    selectedPeople.map((items) => {
      alluser.push(items.key);
    });

    try {
      await list.items.getById(selectedEmployee.ID).update({
        Title: selectedEmployee.Title,
        EmployeeNameId: { results: alluser },
      });

      // alert("Updated Successfully");
      setOpenDialog(false)
      setErrorMessage('');
      const updatedItem = await list.items.getById(selectedEmployee.ID).get();
      setSelectedEmployee((prevEmployee) => ({
        ...prevEmployee!,
        Title: updatedItem && updatedItem.Title ? updatedItem.Title : '',
        Ename: updatedItem && updatedItem.Ename ? updatedItem.Ename : ''
      }));

      setIsPanelOpen(false);
      setIsEditMode(false);
      await fetchData();
    } catch (error) {
      console.error("Error updating data:", error);
    }
  };


  const openCreatePanel = (employee?: IEmployee): void => {
    if (employee) {
      const selectedPeopleFromEmployee: IPersonaProps[] = [
        {
          key: employee.EmployeeName || '',
          text: employee.Ename || '',
        },
      ];
      setSelectedPeople(selectedPeopleFromEmployee);
    } else {
      setSelectedPeople([]);
    }
    setSelectedEmployee(employee || null);
    setIsPanelOpen(true);
    setIsEditMode(!!employee);
    setErrorMessage('');
  };

  const createOrUpdateData = async (): Promise<void> => {
    if (isEditMode) {
      await updateData();
    } else {
      await createData();
    }
  };

  const deleteData = async (employeeId: number): Promise<void> => {
    const web = new Web(webURL);
    await web.lists.getByTitle("Example").items.getById(employeeId).delete();
    alert("Deleted Successfully");
    fetchData().then((data) => console.log(data)).catch((error) => console.log(error));
  };

  const onResolveSuggestions = async (filter: string, selectedItems?: IPersonaProps[]): Promise<IPersonaProps[]> => {
    try {
      const web = new Web(webURL);
      const userList = await web.siteUsers.get();

      const filteredUsers: IPersonaProps[] = userList
        .filter((user: any) => user.Title.toLowerCase().includes(filter.toLowerCase()))
        .map((user: any) => ({
          key: user.Id.toString(),
          text: user.Title,
          secondaryText: user.Email,
        }));

      return filteredUsers;
    } catch (error) {
      console.error("Error fetching user suggestions:", error);
      return [];
    }
  };

  const onPeoplePickerChange = (items?: IPersonaProps[]) => {
    setSelectedPeople(items || []);
    setErrorMessage('');
  };
  // console.log(selectedPeople, "selectedPeople")

  return (
    <div>
      <h1 className={styles.heading}>CRUD Operations - 2</h1>
      <div style={{ display: "flex", gap: '10px' }}>
        <PrimaryButton className={styles.heading} iconProps={addIcon} text="Create" onClick={() => openCreatePanel()} />

      </div>

      <Panel
        isOpen={isPanelOpen}
        headerText={isEditMode ? "Update Employee" : "Create Employee"}
        type={PanelType.custom}
        customWidth='40%'
        closeButtonAriaLabel="Close"
        onDismiss={() => setIsPanelOpen(false)}
      >
        <form>
          <div>
            <Label>Desigantion</Label>
            <TextField
              value={selectedEmployee?.Title || ''}
              onChange={(ev, value) => setSelectedEmployee((prevEmployee) => ({ ...prevEmployee!, Title: value || '' }))}
            />
          </div>

          <div>
            <Label>User</Label>
            <NormalPeoplePicker
              onResolveSuggestions={onResolveSuggestions}
              getTextFromItem={(persona: IPersonaProps) => persona.text || ''}
              pickerSuggestionsProps={{
                suggestionsHeaderText: 'Suggested People',
                noResultsFoundText: 'No results found',
              }}
              onChange={onPeoplePickerChange}
              selectedItems={selectedPeople}
            />
            {errorMessage && <div style={{ color: 'red' }}>{errorMessage}</div>}
          </div>
          <div style={{ display: "flex", textAlign: 'center', gap: '5px' }}>
            <PrimaryButton className={styles.btn} text={isEditMode ? "Update" : "Create"} onClick={createOrUpdateData} />
          </div> </form>
      </Panel>

      <Dialog
        hidden={openDialog}
        dialogContentProps={dialogProps}
      >
        <DialogFooter>

          <PrimaryButton onClick={() => setOpenDialog(true)} text='Okay' />

        </DialogFooter>


      </Dialog>


      <div>
        <h2>Sample Picker</h2>
        {/* <table>
          <thead>
            <tr>
              <th>Sl.No</th>
              <th>Title</th>
              <th>User</th>
              <th>Action</th>
            </tr>
          </thead>
          <tbody>
            {employees.map(employee => (
              <tr key={employee.ID}>
                <td>{employee.ID}</td>
                <td>{employee.Title || ''}</td>
                <td>{employee.EmployeeNameId || ''}</td>
                <td>
                  <PrimaryButton text="Update" onClick={() => openCreatePanel(employee)} />
                  <PrimaryButton text="Delete" onClick={() => deleteData(employee.ID)} />
                </td>
              </tr>
            ))}
          </tbody>
        </table> */}

        <DetailsList
          items={employees}
          columns={[
            { key: "Auto ID", name: "Auto-ID", fieldName: "ID", minWidth: 50, maxWidth: 90 },
            { key: "User", name: "User(PPID)", fieldName: "EmployeeNameId", minWidth: 50, maxWidth: 90 },
            { key: "Name", name: "Name(PP)", fieldName: "Ename", minWidth: 50, maxWidth: 90 },
            { key: "Designation", name: "Designation", fieldName: "Title", minWidth: 50, maxWidth: 90 },
            {
              key: "Actions",
              name: "Actions",
              minWidth: 50,
              onRender: (row) => (
                <div>
                  <IconButton iconProps={editIcon} onClick={() => openCreatePanel(row)} />
                  <IconButton iconProps={deleteIcon} onClick={() => deleteData(row.ID)} />
                </div>
              )
            }
          ]}
        />
      </div>
    </div>
  );
};
export default CRUDtwo;
