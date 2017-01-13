Imports System.IO
Imports LibGit2Sharp
Imports LibGit2Sharp.Handlers
Imports Microsoft.SqlServer.Management.Common
Imports Microsoft.SqlServer.Management.Smo
Imports Newtonsoft.Json

Module Module1

    Public ConfigSetting As ConfigSettingDefinition = New ConfigSettingDefinition

    Class DatabaseDefinition
        Public Name As String = ""
    End Class
	
    Class SQLServerDefinition
        Public Name As String = ""
        Public User As String = ""
        Public Password As String = ""
        Public Databases As List(Of DatabaseDefinition)

        Public Sub New()
            Databases = New List(Of DatabaseDefinition)
        End Sub
    End Class
	
    Class ConfigSettingDefinition
        Public RepositoryWorkingDirectory As String = ""
        Public CommitAuthorName As String = ""
        Public CommitAuthorEmail As String = ""
        Public GitHubUserName As String = ""
        Public GitHubPassword As String = ""
        Public SQLServers As List(Of SQLServerDefinition)

        Public Sub New()
            SQLServers = New List(Of SQLServerDefinition)
        End Sub
    End Class

    Sub Main()

        Dim PathConfigFile = Path.Combine(My.Application.Info.DirectoryPath, "config.json")

        ConfigSetting = LoadConfigSettingFromFile(PathConfigFile)

        Dim repo As Repository = New Repository(ConfigSetting.RepositoryWorkingDirectory)

        For Each Server In ConfigSetting.SQLServers

            Dim conn As New ServerConnection(Server.Name, Server.User, Server.Password) ' -- set SQL server connection given the server name, user name and password
            Dim oSQLServer As New Server(conn)

            Dim ParentFolder = Path.Combine(ConfigSetting.RepositoryWorkingDirectory, Server.Name + "\Jobs")
            My.Computer.FileSystem.CreateDirectory(ParentFolder)

            Dim AllJobs As Dictionary(Of String, String) = New Dictionary(Of String, String)

            For Each Job As Agent.Job In oSQLServer.JobServer.Jobs

                Dim JobDescriptionStr = ""
                Dim JobDescription = Job.Script()

                For Each StrLine In JobDescription
                    JobDescriptionStr = JobDescriptionStr + IIf(String.IsNullOrEmpty(JobDescriptionStr), "", vbNewLine) + StrLine
                Next

                Dim FileNameShort = Job.JobID.ToString + ".sql"

                Dim FileName = Path.Combine(ParentFolder, FileNameShort)

                My.Computer.FileSystem.WriteAllText(FileName, JobDescriptionStr, False)

                AllJobs.Add(FileNameShort, Job.Name)

                'Commit change
                CommitChanges(repo, "Job: " + Job.Name)

            Next

            Dim AllFiles = My.Computer.FileSystem.GetFiles(ParentFolder)
            For Each File In AllFiles

                Dim FI = My.Computer.FileSystem.GetFileInfo(File)

                If Not AllJobs.ContainsKey(FI.Name) Then
                    My.Computer.FileSystem.DeleteFile(File)
                End If

            Next

            'commit deletion
            CommitChanges(repo, "Deleted jobs")

            oSQLServer = Nothing

        Next

        'PUSH changes to Git
        Dim Remote As Remote = repo.Network.Remotes("origin")

        Dim options = New PushOptions()
        options.CredentialsProvider = Function(_url, _user, _cred) New UsernamePasswordCredentials() With {.Username = ConfigSetting.GitHubUserName, .Password = ConfigSetting.GitHubPassword}
        repo.Network.Push(Remote, "refs/heads/master", options)


    End Sub

    Sub CommitChanges(repo As Repository, Message As String)

        Try

            Commands.Stage(repo, "*")

            Dim author As Signature = New Signature(ConfigSetting.CommitAuthorName, ConfigSetting.CommitAuthorEmail, DateTime.Now)
            Dim committer As Signature = author

            Dim commit As Commit = repo.Commit(Message, author, committer)

        Catch ex As Exception

        End Try

    End Sub

    Public Function LoadConfigSettingFromFile(ConfigFilePath As String) As ConfigSettingDefinition

        If My.Computer.FileSystem.FileExists(ConfigFilePath) Then

            Dim JsonText = My.Computer.FileSystem.ReadAllText(ConfigFilePath)

            Dim ConfigSettingObj = JsonConvert.DeserializeObject(Of ConfigSettingDefinition)(JsonText)

            Return ConfigSettingObj

        End If

        Return New ConfigSettingDefinition

    End Function


End Module
