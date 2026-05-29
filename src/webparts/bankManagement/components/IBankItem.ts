export interface IBankItem {
  Id?: number;
  CustomerName: string;
  Email: string;
  PhoneNumber: number;
  AccountType: string;
  Balance: number;
  IsActive: boolean;
  Branch: string;
  DateOfJoining: string;
}

export interface IMasterItem {
  Id?: number;
  BranchCode: number;
  BranchName: string;
  BranchLocation: string;
}
