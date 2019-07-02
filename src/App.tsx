import React, { Component } from 'react';
import { BrowserRouter as Router, Route } from 'react-router-dom';
import { Container } from 'reactstrap';
import NavBar from './components/NavBar';
import ErrorMessage from './components/ErrorMessage';
import WelcomePage from './pages/WelcomePage';
import config from './Config';
import { UserAgentApplication } from 'msal';
import { getUserDetails } from './services/GraphService';
import 'bootstrap/dist/css/bootstrap.css';
import Calendar from './pages/Calendar';
import OneDrive from './pages/OneDrive';
import OneDriveList from './pages/OneDriveList';
import OneDriveTable from './pages/OneDriveTable';

class App extends Component<any, any> {
  private userAgentApplication;

  constructor(props) {
    super(props);
    this.userAgentApplication = new UserAgentApplication({
      auth: {
        clientId: config.appId
      },
      cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: true
      }
    });

    var user = this.userAgentApplication.getAccount();

    this.state = {
      isAuthenticated: user !== null,
      user: {},
      error: null
    };

    if (user) {
      // Enhance user object with data from Graph
      this.getUserProfile();
    }
  }
  async login() {
    try {
      await this.userAgentApplication.loginPopup({
        scopes: config.scopes,
        prompt: 'select_account'
      });
      await this.getUserProfile();
    } catch (err) {
      var errParts = err.split('|');
      this.setState({
        isAuthenticated: false,
        user: {},
        error: { message: errParts[1], debug: errParts[0] }
      });
    }
  }
  logout() {
    this.userAgentApplication.logout();
  }
  async getUserProfile() {
    try {
      // Get the access token silently
      // If the cache contains a non-expired token, this function
      // will just return the cached token. Otherwise, it will
      // make a request to the Azure OAuth endpoint to get a token

      var accessToken = await this.userAgentApplication.acquireTokenSilent({
        scopes: config.scopes
      });

      if (accessToken) {
        // Get the user's profile from Graph
        var user = await getUserDetails();
        this.setState({
          isAuthenticated: true,
          user: {
            displayName: user.displayName,
            email: user.mail || user.userPrincipalName
          },
          error: null
        });
      }
    } catch (err) {
      var error = {};
      if (typeof err === 'string') {
        var errParts = err.split('|');
        error =
          errParts.length > 1 ? { message: errParts[1], debug: errParts[0] } : { message: err };
      } else {
        error = {
          message: err.message,
          debug: JSON.stringify(err)
        };
      }

      this.setState({
        isAuthenticated: false,
        user: {},
        error: error
      });
    }
  }

  render() {
    let error;
    if (this.state.error) {
      error = <ErrorMessage message={this.state.error.message} debug={this.state.error.debug} />;
    }

    return (
      <Router>
        <div>
          <NavBar
            isAuthenticated={this.state.isAuthenticated}
            authButtonMethod={
              this.state.isAuthenticated ? this.logout.bind(this) : this.login.bind(this)
            }
            user={this.state.user}
          />
          <Container>
            {error}
            <Route
              exact
              path="/"
              render={props => (
                <WelcomePage
                  {...props}
                  isAuthenticated={this.state.isAuthenticated}
                  user={this.state.user}
                  authButtonMethod={this.login.bind(this)}
                />
              )}
            />
            <Route
              exact
              path="/calendar"
              render={props => <Calendar {...props} showError={this.setErrorMessage.bind(this)} />}
            />
            <Route
              exact
              path="/onedrivelist"
              render={props => (
                <OneDriveList {...props} showError={this.setErrorMessage.bind(this)} />
              )}
            />
            <Route
              exact
              path="/onedrivelist/:documentId"
              render={props => <OneDrive {...props} showError={this.setErrorMessage.bind(this)} />}
            />
            <Route
              exact
              path="/table/:documentId/:sheetName"
              render={props => (
                <OneDriveTable {...props} showError={this.setErrorMessage.bind(this)} />
              )}
            />
          </Container>
        </div>
      </Router>
    );
  }

  setErrorMessage(message, debug) {
    this.setState({
      error: { message: message, debug: debug }
    });
  }
}

export default App;
